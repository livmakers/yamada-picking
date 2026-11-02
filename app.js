'use strict';
/* ヤマダ 出荷ピッキング(機能B) — ブラウザだけで動く / 通信なしで照合 / 月額ゼロ */
(function(){
  const $ = s => document.querySelector(s);
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const half = s => String(s == null ? '' : s).replace(/[Ａ-Ｚａ-ｚ０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0)).replace(/　/g, ' ').trim();
  const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  const p2 = n => (n < 10 ? '0' : '') + n;
  const fmt = ts => { const d = new Date(ts); return d.getFullYear() + '-' + p2(d.getMonth()+1) + '-' + p2(d.getDate()) + ' ' + p2(d.getHours()) + ':' + p2(d.getMinutes()) + ':' + p2(d.getSeconds()); };

  // ================= 保存(スマホの中のバインダー) =================
  const DB = {
    db: null, mem: { kv: {}, log: {} },
    open() { return new Promise(res => { try {
      const r = indexedDB.open('yamada-picking', 1);
      r.onupgradeneeded = () => { r.result.createObjectStore('kv'); r.result.createObjectStore('log', { keyPath: 'id' }); };
      r.onsuccess = () => { DB.db = r.result; res(); }; r.onerror = () => res();
    } catch (e) { res(); } }); },
    get(k) { return new Promise(res => { if (!DB.db) return res(DB.mem.kv[k]); const q = DB.db.transaction('kv').objectStore('kv').get(k); q.onsuccess = () => res(q.result); q.onerror = () => res(); }); },
    set(k, v) { return new Promise(res => { if (!DB.db) { DB.mem.kv[k] = v; return res(); } const t = DB.db.transaction('kv', 'readwrite'); t.objectStore('kv').put(v, k); t.oncomplete = res; t.onerror = res; }); },
    putLog(e) { return new Promise(res => { if (!DB.db) { DB.mem.log[e.id] = e; return res(); } const t = DB.db.transaction('log', 'readwrite'); t.objectStore('log').put(e); t.oncomplete = res; t.onerror = res; }); },
    allLog() { return new Promise(res => { if (!DB.db) return res(Object.values(DB.mem.log)); const q = DB.db.transaction('log').objectStore('log').getAll(); q.onsuccess = () => res(q.result || []); q.onerror = () => res([]); }); }
  };

  // ================= 状態 =================
  const DEF_MASTER = [['4559','LL','CHA'],['4559','L','BRN'],['4559','L','CHA'],['4559','M','BEG'],['4559','M','WHT'],['4558','LL','BRN'],['4558','L','BRN'],['4558','L','CHA'],['4558','L','BEG']];
  const DEF_COLORS = { CHA:'チャコール', BRN:'ブラウン', BEG:'ベージュ', WHT:'ホワイト' };
  const SAMPLE_ORDERS = '店名,品番,サイズ,色,数量\nテストA店,4559,LL,CHA,3\nテストA店,4559,L,BRN,2\nテストA店,4558,L,CHA,4\nテストA店,4559,M,BEG,1\nテストB店,4559,LL,CHA,2\nテストB店,4558,LL,BRN,5\nテストB店,4559,M,WHT,3\nテストB店,4558,L,BEG,2\nテストC店,4559,L,CHA,6\nテストC店,4558,L,BRN,2\nテストC店,4559,L,BRN,1\nテストC店,4558,L,CHA,1';
  const S = {
    master: [], colors: {}, colmap: { store: 1, hin: 2, size: 3, col: 4, qty: 5, header: 'auto' },
    orders: null, worker: '', assigned: [], picked: {}, syncUrl: '', syncLog: [], log: [],
    selStore: '', single: false, view: 'home', exp: {}, lastBackup: null
  };
  const persist = k => DB.set(k, S[k]);

  const colorName = c => S.colors[c] || c;
  const keyOf = (h, s, c) => 'YMD-' + h + '-' + s + '-' + c;
  const parseKey = k => { const m = /^YMD-([A-Za-z0-9]+)-([A-Za-z0-9]+)-([A-Za-z0-9]+)$/.exec(k); return m ? { hin: m[1], size: m[2].toUpperCase(), col: m[3].toUpperCase() } : null; };
  const nameOf = k => { const p = parseKey(k); return p ? p.hin + ' ' + p.size + ' ' + colorName(p.col) : k; };
  const masterHas = k => S.master.some(m => keyOf(m.hin, m.size, m.col) === k);

  // ================= CSV / Excel =================
  function decode(buf) { try { return new TextDecoder('utf-8', { fatal: true }).decode(buf); } catch (e) { return new TextDecoder('shift_jis').decode(buf); } }
  function parseCSV(text) {
    text = text.replace(/^﻿/, '');
    const out = []; let row = [], cur = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (q) { if (c === '"') { if (text[i+1] === '"') { cur += '"'; i++; } else q = false; } else cur += c; }
      else if (c === '"') q = true;
      else if (c === ',') { row.push(cur); cur = ''; }
      else if (c === '\n' || c === '\r') { if (c === '\r' && text[i+1] === '\n') i++; row.push(cur); cur = ''; out.push(row); row = []; }
      else cur += c;
    }
    if (cur !== '' || row.length) { row.push(cur); out.push(row); }
    return out.filter(r => r.some(x => String(x).trim() !== ''));
  }
  function loadRows(file) {
    return new Promise((res, rej) => {
      const fr = new FileReader();
      fr.onload = () => {
        try {
          if (/\.xlsx?$/i.test(file.name)) {
            if (!window.XLSX) return rej(new Error('Excel読み込み部品がありません。CSVで保存し直してください'));
            const wb = XLSX.read(fr.result, { type: 'array' });
            res(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: false, defval: '' }));
          } else res(parseCSV(decode(fr.result)));
        } catch (e) { rej(e); }
      };
      fr.onerror = () => rej(new Error('ファイルを読めません'));
      fr.readAsArrayBuffer(file);
    });
  }
  function toCSV(rows) { return rows.map(r => r.map(v => { v = v == null ? '' : String(v); return /[",\n\r]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; }).join(',')).join('\r\n'); }
  function download(name, text) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob(['﻿' + text], { type: 'text/csv;charset=utf-8' }));
    a.download = name; document.body.appendChild(a); a.click(); a.remove();
  }

  // ================= 受注の取り込み =================
  function colorCode(raw) {
    const up = raw.toUpperCase();
    if (S.colors[up]) return up;
    const hit = Object.keys(S.colors).find(c => S.colors[c] === raw);
    return hit || '';
  }
  function buildOrders(rows, src) {
    const cm = S.colmap, errs = [], need = {}, stores = [];
    let start = 0, ok = 0, total = 0;
    if (cm.header === 'yes') start = 1;
    else if (cm.header === 'auto' && rows[0] && !/^\d+$/.test(half(rows[0][cm.qty - 1]))) start = 1;
    for (let i = start; i < rows.length; i++) {
      const r = rows[i], line = i + 1, pr = [];
      const store = String(r[cm.store - 1] == null ? '' : r[cm.store - 1]).trim();
      const hin = half(r[cm.hin - 1]), size = half(r[cm.size - 1]).toUpperCase(), cRaw = half(r[cm.col - 1]);
      const qty = parseInt(half(r[cm.qty - 1]), 10);
      const col = colorCode(cRaw);
      if (!store) pr.push('店名が空'); if (!hin) pr.push('品番が空'); if (!size) pr.push('サイズが空');
      if (!col) pr.push('色「' + cRaw + '」が色コード表にない');
      if (!(qty >= 1)) pr.push('数量が数字でない');
      const key = keyOf(hin, size, col);
      if (!pr.length && S.master.length && !masterHas(key)) pr.push(nameOf(key) + ' は商品マスターにない');
      if (pr.length) { errs.push(line + '行目: ' + pr.join('、')); continue; }
      if (!need[store]) { need[store] = {}; stores.push(store); }
      need[store][key] = (need[store][key] || 0) + qty; ok++; total += qty;
    }
    return { batch: uid(), loadedAt: Date.now(), file: src, need, stores, total, rowsOk: ok, errs: errs.slice(0, 100), errCount: errs.length };
  }
  async function setOrders(o) {
    S.orders = o; S.picked = {}; S.assigned = S.assigned.filter(s => o.stores.indexOf(s) >= 0);
    await persist('orders'); await persist('picked'); await persist('assigned');
  }

  // ================= 進捗の計算 =================
  function doneMap() {
    const m = {};
    if (!S.orders) return m;
    S.log.forEach(e => { if (e.batch === S.orders.batch && e.kind === 'ok' && !e.undone) { const k = e.store + '|' + e.key; m[k] = (m[k] || 0) + e.qty; } });
    return m;
  }
  function candidates(key) {
    const dm = doneMap(), out = [];
    S.orders.stores.forEach(s => {
      const n = (S.orders.need[s] || {})[key] || 0;
      if (n > 0) { const d = dm[s + '|' + key] || 0; out.push({ store: s, need: n, done: d, remain: n - d }); }
    });
    return out;
  }
  function storeStat(s) {
    const dm = doneMap(), need = S.orders.need[s] || {}; let n = 0, d = 0;
    Object.keys(need).forEach(k => { n += need[k]; d += Math.min(need[k], dm[s + '|' + k] || 0); });
    return { need: n, done: d };
  }
  const visibleStores = () => (S.assigned.length ? S.orders.stores.filter(s => S.assigned.indexOf(s) >= 0) : S.orders.stores);

  // ================= 記録 =================
  function addLog(e) {
    e.id = uid(); e.ts = Date.now(); e.worker = S.worker; e.batch = S.orders ? S.orders.batch : ''; e.synced = false; e.undone = false;
    S.log.push(e); DB.putLog(e); return e;
  }

  // ================= 音・ブルブル =================
  const Snd = {
    ctx: null,
    init() { try { if (!Snd.ctx) Snd.ctx = new (window.AudioContext || window.webkitAudioContext)(); if (Snd.ctx.state === 'suspended') Snd.ctx.resume(); } catch (e) {} },
    beep(f, d, type, at) { try { if (!Snd.ctx) return; const o = Snd.ctx.createOscillator(), g = Snd.ctx.createGain(), t = Snd.ctx.currentTime + (at || 0);
      o.type = type || 'sine'; o.frequency.value = f; g.gain.setValueAtTime(0.4, t); g.gain.exponentialRampToValueAtTime(0.001, t + d);
      o.connect(g); g.connect(Snd.ctx.destination); o.start(t); o.stop(t + d + 0.02); } catch (e) {} },
    ok() { Snd.beep(1000, 0.12); try { navigator.vibrate && navigator.vibrate(60); } catch (e) {} },
    done() { Snd.beep(784, 0.12); Snd.beep(1047, 0.12, 'sine', 0.14); Snd.beep(1319, 0.2, 'sine', 0.28); },
    err() { Snd.beep(200, 0.3, 'square'); Snd.beep(160, 0.4, 'square', 0.34); try { navigator.vibrate && navigator.vibrate([200, 80, 200]); } catch (e) {} }
  };

  // ================= カメラでQR読み取り =================
  const Scanner = {
    stream: null, timer: null, video: null, cb: null, last: '', lastT: 0,
    start(video, cb) {
      Scanner.video = video; Scanner.cb = cb;
      if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) return Promise.reject(new Error('nocamera'));
      return navigator.mediaDevices.getUserMedia({ video: { facingMode: { ideal: 'environment' }, width: { ideal: 1280 }, height: { ideal: 720 } }, audio: false })
        .then(st => { Scanner.stream = st; video.srcObject = st; video.setAttribute('playsinline', ''); return video.play(); })
        .then(() => Scanner.loop());
    },
    loop() {
      const c = document.createElement('canvas'), g = c.getContext('2d', { willReadFrequently: true });
      const tick = () => {
        if (!Scanner.stream) return;
        const v = Scanner.video;
        if (v.readyState >= 2 && v.videoWidth) {
          const w = Math.min(640, v.videoWidth), h = Math.round(v.videoHeight * w / v.videoWidth);
          c.width = w; c.height = h; g.drawImage(v, 0, 0, w, h);
          const r = jsQR(g.getImageData(0, 0, w, h).data, w, h, { inversionAttempts: 'dontInvert' });
          if (r && r.data) {
            const now = Date.now();
            if (r.data !== Scanner.last || now - Scanner.lastT > 1500) { Scanner.last = r.data; Scanner.lastT = now; Scanner.cb(r.data); } else Scanner.lastT = now;
          }
        }
        Scanner.timer = setTimeout(tick, 110);
      };
      tick();
    },
    stop() { clearTimeout(Scanner.timer); if (Scanner.stream) Scanner.stream.getTracks().forEach(t => t.stop()); Scanner.stream = null; }
  };

  // ================= 画面の小道具 =================
  function toast(msg, ms) { const t = $('#toast'); t.textContent = msg; t.hidden = false; clearTimeout(toast.t); toast.t = setTimeout(() => { t.hidden = true; }, ms || 3200); }
  let flashT = null;
  function flash(type, msg, ms) {
    const f = $('#flash'); if (!f) return;
    f.className = type; f.innerHTML = msg; f.hidden = false; clearTimeout(flashT);
    if (ms) flashT = setTimeout(() => { f.hidden = true; }, ms);
  }
  function updNet() {
    const pend = S.log.filter(e => !e.synced).length;
    $('#net').textContent = (navigator.onLine ? '● オンライン' : '○ 電波なし') + (S.syncUrl && pend ? '  未送信' + pend : ''); updBk();
  }
  function needWorker() { if (!S.worker.trim()) { toast('先に「お名前」を入力してください'); const i = $('#inName'); if (i) i.focus(); return false; } return true; }
  function needOrders() { if (!S.orders) { toast('先に「受注を取り込む」をしてください'); return false; } return true; }
  const bar = (d, n) => '<div class="bar"><i style="width:' + (n ? Math.round(d * 100 / n) : 0) + '%"></i></div>';

  // ================= 各画面 =================
  const V = {};
  let cleanup = null;
  function go(name) {
    if (cleanup) { cleanup(); cleanup = null; }
    S.view = name; const v = V[name];
    $('#view').innerHTML = v.html(); $('#title').textContent = v.title;
    $('#btnHome').hidden = (name === 'home'); $('#flash').hidden = true; window.scrollTo(0, 0);
    if (v.init) cleanup = v.init() || null;
    updNet();
  }

  // ---- ホーム ----
  V.home = {
    title: 'ヤマダ 出荷ピッキング',
    html() {
      const o = S.orders; let st = '受注: まだ取り込んでいません';
      if (o) st = '受注: ' + o.stores.length + '店 / 合計' + o.total + '枚(' + esc(o.file) + ' ' + fmt(o.loadedAt).slice(5, 16) + ')';
      return '<div class="card"><label class="fl">あなたのお名前(記録に残ります)</label><input type="text" id="inName" data-chg="name" value="' + esc(S.worker) + '" placeholder="例: 山田"></div>' +
        '<div class="status" style="margin:6px 4px">' + st + '</div>' +
        '<button class="btn gray" data-act="go" data-v="check">⓪ 朝の点検(始業前に押す)</button>' +
        '<button class="btn gray" data-act="go" data-v="import">① 受注を取り込む</button>' +
        '<button class="btn" data-act="go" data-v="pick">② 総量ピッキング(まとめて集める)</button>' +
        '<button class="btn green" data-act="go" data-v="sort">③ 仕分け(QRを読んで店ごとに分ける)</button>' +
        '<button class="btn gray" data-act="go" data-v="progress">④ 進捗(店ごとの完了・未完了)</button>' +
        '<button class="btn gray" data-act="go" data-v="output">⑤ 出荷明細・作業記録を出す</button>' +
        '<div class="card"><b>もしもの時</b><p class="hint">スマホが壊れた・電池切れ・QRが読めない時の備えです。</p>' +
        '<button class="btn gray small" data-act="backup">バックアップを保存・共有(LINE等へ)</button>' +
        '<label class="btn gray small">引き継ぎ(バックアップを読み込む)<input type="file" accept=".json,application/json" data-file="backup" hidden></label>' +
        '<button class="btn gray small" data-act="paper"' + (S.orders ? '' : ' disabled') + '>紙の予備を印刷(総量ピッキング表+手書き用出荷明細)</button></div>' +
        '<button class="btn gray small" data-act="go" data-v="settings">設定(担当の店・記録の送信先)</button>' +
        '<button class="btn gray small" disabled>棚卸し(機能C・準備中)</button>';
    }
  };

  // ---- 取り込み ----
  V.import = {
    title: '受注を取り込む',
    html() {
      const o = S.orders, cm = S.colmap;
      const numIn = (k, l) => '<div><label class="fl">' + l + '</label><input type="number" min="1" data-chg="colmap" data-k="' + k + '" value="' + cm[k] + '"></div>';
      let res = '';
      if (o) {
        res = '<div class="card"><div class="big ok">' + o.stores.length + '店 / ' + o.rowsOk + '行 / 合計' + o.total + '枚</div><div class="status">' + esc(o.file) + '</div>' +
          (o.errCount ? '<div class="ng" style="margin-top:8px;font-size:17px"><b>取り込めなかった行が' + o.errCount + '件あります</b><br>' + o.errs.map(esc).join('<br>') + '</div>' : '<div class="status ok">エラーはありません</div>') + '</div>';
      }
      return '<button class="btn green" data-act="sample">まず試す(テスト9点+ダミー3店)</button>' + res +
        (o ? '<button class="btn gray small" data-act="paper">紙の予備を印刷(総量ピッキング表+手書き用出荷明細)</button>' : '') +
        '<div class="card"><b>商品マスター</b><div class="status">' + (S.master.length ? S.master.length + '品 読込済(QR照合に使います)' : '未読込(読み込むと、未登録の商品QRも警告できます)') + '</div>' +
        '<label class="btn gray small">CSVを選ぶ<input type="file" accept=".csv,text/csv" data-file="master" hidden></label></div>' +
        '<div class="card"><b>色コード表</b><div class="status">' + Object.keys(S.colors).length + '色 読込済</div>' +
        '<label class="btn gray small">CSVを選ぶ<input type="file" accept=".csv,text/csv" data-file="colors" hidden></label></div>' +
        '<div class="card"><b>受注ファイル(CSV または Excel)</b>' +
        '<p class="hint">列の番号を合わせてください(左端=1)。量販店の形式が変わったら、ここの数字を変えるだけです。</p>' +
        '<div class="row">' + numIn('store', '店名の列') + numIn('hin', '品番の列') + numIn('size', 'サイズの列') + '</div>' +
        '<div class="row">' + numIn('col', '色の列') + numIn('qty', '数量の列') +
        '<div><label class="fl">1行目</label><select data-chg="colmap" data-k="header"><option value="auto"' + (cm.header === 'auto' ? ' selected' : '') + '>自動判定</option><option value="yes"' + (cm.header === 'yes' ? ' selected' : '') + '>見出し行</option><option value="no"' + (cm.header === 'no' ? ' selected' : '') + '>データ行</option></select></div></div>' +
        '<label class="btn small" style="margin-top:14px">受注ファイルを選ぶ<input type="file" accept=".csv,.xlsx,.xls,text/csv" data-file="orders" hidden></label>' +
        '<p class="hint">新しい受注を取り込むと、仕分けの進捗は0から始まります(これまでの作業記録は残ります)。</p></div>';
    }
  };

  // ---- 総量ピッキング ----
  V.pick = {
    title: '総量ピッキング',
    html() {
      if (!S.orders) return '<p class="hint">受注が未取込です。</p>';
      const tot = {};
      S.orders.stores.forEach(s => Object.keys(S.orders.need[s]).forEach(k => { tot[k] = (tot[k] || 0) + S.orders.need[s][k]; }));
      const keys = Object.keys(tot).sort(); const got = keys.filter(k => S.picked[k]).length;
      return '<div class="card"><div class="big">集めた ' + got + ' / ' + keys.length + ' 品</div>' + bar(got, keys.length) + '</div>' +
        keys.map(k => '<div class="card tick' + (S.picked[k] ? ' done' : '') + '" data-act="ptoggle" data-k="' + k + '"><div class="box">' + (S.picked[k] ? '✓' : '') + '</div>' +
          '<div class="name">' + esc(nameOf(k)) + '</div><div class="num">' + tot[k] + '枚</div></div>').join('') +
        '<p class="hint">集め終えた商品をタップ(✓)。「合計○枚」を倉庫から集めてから、③仕分けへ進みます。</p>';
    }
  };

  // ---- 仕分け(種まき) ----
  let P = null;           // 読み取った商品の確認カード
  let listening = false;
  V.sort = {
    title: '仕分け',
    html() {
      if (!S.orders) return '<p class="hint">受注が未取込です。</p>';
      const opts = '<option value="">店を指定しない(商品から店を探す)</option>' + visibleStores().map(s => '<option value="' + esc(s) + '"' + (S.selStore === s ? ' selected' : '') + '>' + esc(s) + '</option>').join('');
      return '<div class="video"><video id="cam" muted playsinline></video><div class="frame"></div></div>' +
        '<div id="camMsg" class="status"></div>' +
        '<div class="row" style="margin:8px 0"><select id="selStore" data-chg="selstore">' + opts + '</select></div>' +
        '<label class="status" style="display:flex;gap:10px;align-items:center;margin:6px 0"><input type="checkbox" data-chg="single" style="width:28px;height:28px"' + (S.single ? ' checked' : '') + '> 1個ずつモード(読むたびに1枚で自動確定)</label>' +
        '<div id="panel"></div>' +
        '<details class="card"><summary style="font-size:18px">カメラが使えない時 / バーコードリーダー用</summary><label class="fl">QRの文字を入力</label><input type="text" id="manual" placeholder="YMD-4559-LL-CHA" autocapitalize="characters"><button class="btn small" data-act="manual">読み取り</button></details>';
    },
    init() {
      if (!S.orders) return;
      P = null; renderPanel();
      Snd.init(); keepAwake(true);
      const v = $('#cam');
      Scanner.start(v, onCode).catch(e => { const m = $('#camMsg'); if (m) m.innerHTML = '<span class="ng"><b>カメラを使えません。</b>ブラウザの「カメラを許可」を確認してください。下の入力欄でも試せます。</span>'; });
      return () => { Scanner.stop(); P = null; keepAwake(false); };
    }
  };

  function renderPanel() {
    const el = $('#panel'); if (!el) return;
    const vd = document.querySelector('.video'); if (vd) vd.style.display = P ? 'none' : '';   // 確認カード表示中はカメラ枠を畳み、確定ボタンを画面内に収める
    if (!P) { el.innerHTML = '<div class="card" style="text-align:center"><div class="big">QRをかざしてください</div></div>' + manualList() + lastLine(); return; }
    const c = P.list.find(x => x.store === P.store) || P.list[0];
    const chips = P.list.length > 1 ? '<div class="status">この商品が必要な店(タップで切り替え)</div><div class="chips">' +
      P.list.map(x => '<button class="chip' + (x.store === P.store ? ' on' : '') + '" data-act="pickstore" data-s="' + esc(x.store) + '">' + esc(x.store) + ' 残' + x.remain + '</button>').join('') + '</div>' : '';
    const pads = [1,2,3,4,5,6,7,8,9,10].map(n => '<button data-act="setq" data-n="' + n + '"' + (P.qty === n ? ' class="on"' : '') + '>' + n + '</button>').join('');
    el.innerHTML = '<div class="card"><div class="big">' + esc(nameOf(P.key)) + '</div>' +
      '<div class="huge ok">→ ' + esc(c.store) + ' へ</div><div class="status">この店の注文 ' + c.need + '枚 / 済 ' + c.done + '枚 / <b>残り ' + c.remain + '枚</b></div>' + chips +
      '<div class="qty">' + P.qty + '<span style="font-size:28px">枚</span></div>' +
      '<div class="pad">' + pads + '</div>' +
      '<div class="row"><button class="btn gray small" data-act="decq">− 1</button><button class="btn gray small" data-act="incq">＋ 1</button><button class="btn gray small" data-act="allq">残り全部(' + c.remain + ')</button></div>' +
      '<button class="btn gray small" data-act="voice">' + (listening ? '🎤 聞いています…(数字を言う)' : '🎤 声で数字を入れる') + '</button>' +
      (P.heard ? '<div class="status">聞き取り: 「' + esc(P.heard) + '」</div>' : '') +
      '<button class="btn green" data-act="confirm">確定(' + P.qty + '枚を ' + esc(c.store) + ' へ)</button>' +
      '<button class="btn gray small" data-act="cancel">やめる</button></div>' +
      lastLine();
  }
  function lastLine() {
    const e = S.log.filter(x => x.kind === 'ok' && !x.undone && x.batch === S.orders.batch).slice(-1)[0];
    return e ? '<div class="card status">直前: ' + esc(nameOf(e.key)) + ' → ' + esc(e.store) + ' ' + e.qty + '枚 <button class="chip" data-act="undo">取り消す</button></div>' : '';
  }
  function fail(type, msg, key) {
    Snd.err(); flash('err', msg); addLog({ kind: 'err', errType: type, key: key || '' });
  }
  function onCode(text, via) {
    if (!S.orders || P) return;
    text = String(text).trim();
    if (/^YMD-[SC]\d+$/.test(text)) return fail('badqr', 'これは棚/ケースのQRです。<br>仕分けでは商品のQRを読んでください');
    const q = parseKey(text);
    if (!q) return fail('badqr', 'ヤマダの商品QRではありません');
    const key = keyOf(q.hin, q.size, q.col);
    if (S.master.length && !masterHas(key)) return fail('nomaster', '商品マスターにない商品です<br>' + esc(nameOf(key)), key);
    const all = candidates(key);
    if (!all.length) return fail('notInOrder', '注文にない商品です<br>' + esc(nameOf(key)), key);
    const open = all.filter(x => x.remain > 0);
    if (!open.length) return fail('over', '数量オーバー<br>この商品は全店ぶん そろっています', key);
    let list = open;
    if (S.selStore) {
      const mine = open.find(x => x.store === S.selStore);
      if (!mine) {
        const done = all.find(x => x.store === S.selStore);
        if (done) return fail('over', '数量オーバー<br>' + esc(S.selStore) + 'は この商品 完了ずみ<br>必要な店: ' + open.map(x => esc(x.store)).join('、'), key);
        return fail('wrongStore', '店違い<br>この商品は ' + open.map(x => esc(x.store) + ' ' + x.remain + '枚').join('、') + ' 向けです', key);
      }
      list = [mine];
    } else if (S.assigned.length) {
      list = open.filter(x => S.assigned.indexOf(x.store) >= 0);
      if (!list.length) return fail('wrongStore', '担当外の店の商品です<br>' + open.map(x => esc(x.store)).join('、'), key);
    }
    P = { key, list, store: list[0].store, qty: 1, heard: null, parsed: null, via: via || 'scan' };
    Snd.beep(660, 0.08);
    flash('ok', '読み取りました', 1200);
    if (S.single && list.length === 1) return confirmQty(true);
    renderPanel();
  }
  function confirmQty(auto) {
    if (!P) return;
    const c = candidates(P.key).find(x => x.store === P.store);
    if (!(P.qty >= 1)) return toast('数量を入れてください');
    if (!c || P.qty > c.remain) {
      return fail('over', '数量オーバー<br>' + esc(P.store) + 'は あと' + (c ? c.remain : 0) + '枚 なのに ' + P.qty + '枚', P.key);
    }
    const voiceFixed = P.parsed != null && P.qty !== P.parsed;
    const input = auto ? 'auto1' : (P.parsed != null && P.qty === P.parsed ? 'voice' : 'button');
    const p = parseKey(P.key);
    addLog({ kind: 'ok', key: P.key, hin: p.hin, size: p.size, col: p.col, store: P.store, qty: P.qty, input, via: P.via, voiceHeard: P.heard || '', voiceFixed });
    const st = storeStat(P.store), name = P.store;
    P = null; renderPanel();
    if (st.done >= st.need) { Snd.done(); flash('ok', '◎ ' + esc(name) + ' 全部そろいました!', 4000); }
    else { Snd.ok(); flash('ok', '✔ ' + esc(name) + ' へ記録 (残り ' + (st.need - st.done) + '枚)', 2200); }
    renderPanel(); trySync();
  }

  // ---- 声で数字 ----
  function parseJaNum(s) {
    s = half(s).replace(/[,\s、。]/g, '');
    const m = s.match(/\d+/); if (m) return parseInt(m[0], 10);
    const d = { '〇': 0, '零': 0, '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9 };
    let tot = 0, cur = 0, found = false;
    for (const ch of s) {
      if (d[ch] != null) { cur = d[ch]; found = true; }
      else if (ch === '十') { tot += (cur || 1) * 10; cur = 0; found = true; }
      else if (ch === '百') { tot += (cur || 1) * 100; cur = 0; found = true; }
    }
    return found ? tot + cur : null;
  }
  function startVoice() {
    const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SR) { toast('この端末は音声入力に対応していません。ボタンで入力してください'); addLog({ kind: 'voice_err', errType: 'unsupported' }); return; }
    const r = new SR(); r.lang = 'ja-JP'; r.interimResults = false; r.maxAlternatives = 1;
    listening = true; renderPanel();
    r.onresult = ev => {
      const t = ev.results[0][0].transcript, n = parseJaNum(t);
      if (P) { P.heard = t; if (n != null && n >= 1 && n <= 999) { P.qty = n; P.parsed = n; } else { P.parsed = null; toast('「' + t + '」は数字にできません。ボタンで入れてください'); addLog({ kind: 'voice_err', errType: 'unparsed' }); } }
    };
    r.onerror = ev => { toast('音声が使えません(' + ev.error + ')。ボタンで入力してください'); addLog({ kind: 'voice_err', errType: ev.error || 'error' }); };
    r.onend = () => { listening = false; renderPanel(); };
    try { r.start(); } catch (e) { listening = false; renderPanel(); }
  }

  // ---- 進捗 ----
  V.progress = {
    title: '進捗(店ごと)',
    html() {
      if (!S.orders) return '<p class="hint">受注が未取込です。</p>';
      const dm = doneMap(); let allN = 0, allD = 0;
      const cards = S.orders.stores.map(s => {
        const st = storeStat(s); allN += st.need; allD += st.done;
        const full = st.done >= st.need;
        let det = '';
        if (S.exp[s]) {
          det = '<table><tr><th>商品</th><th class="n">注文</th><th class="n">済</th><th class="n">残</th></tr>' +
            Object.keys(S.orders.need[s]).sort().map(k => { const n = S.orders.need[s][k], d = dm[s + '|' + k] || 0; return '<tr><td>' + esc(nameOf(k)) + '</td><td class="n">' + n + '</td><td class="n">' + d + '</td><td class="n ' + (n - d ? 'ng' : 'ok') + '">' + (n - d) + '</td></tr>'; }).join('') + '</table>';
        }
        return '<div class="card" data-act="pexp" data-s="' + esc(s) + '"><div class="row"><div class="big">' + (full ? '✅ ' : '⏳ ') + esc(s) + '</div><div class="big" style="text-align:right">' + st.done + '/' + st.need + '</div></div>' + bar(st.done, st.need) + det + '</div>';
      }).join('');
      return '<div class="card"><div class="big">全体 ' + allD + ' / ' + allN + ' 枚</div>' + bar(allD, allN) + '</div>' + cards + '<p class="hint">店をタップすると、残りの商品が見えます。</p>';
    }
  };

  // ---- 出力 ----
  function shipRows() {
    const dm = doneMap(), rows = [['店名', '品番', 'サイズ', '色コード', '色名', '注文数', '実績数', '差(実績-注文)']];
    S.orders.stores.forEach(s => Object.keys(S.orders.need[s]).sort().forEach(k => {
      const p = parseKey(k), n = S.orders.need[s][k], d = dm[s + '|' + k] || 0;
      rows.push([s, p.hin, p.size, p.col, colorName(p.col), n, d, d - n]);
    }));
    return rows;
  }
  function metrics() {
    const by = {}, errs = {};
    let vUse = 0, vFix = 0, vFail = 0, manual = 0;
    S.log.forEach(e => {
      if (e.kind === 'ok' && !e.undone && (e.via === 'manual' || e.via === 'typed')) manual++;
      if (e.kind === 'err') errs[e.errType] = (errs[e.errType] || 0) + 1;
      if (e.kind === 'voice_err') vFail++;
      if (e.kind === 'ok' && !e.undone) { if (e.input === 'voice' || e.voiceFixed) vUse++; if (e.voiceFixed) vFix++; }
      const w = e.worker || '(名前なし)'; (by[w] = by[w] || { ts: [], scans: 0, pcs: 0 }).ts.push(e.ts);
      if (e.kind === 'ok' && !e.undone) { by[w].scans++; by[w].pcs += e.qty; }
    });
    const workers = Object.keys(by).map(w => {
      const t = by[w].ts.sort((a, b) => a - b); let act = 0;
      for (let i = 1; i < t.length; i++) { const g = t[i] - t[i-1]; if (g <= 300000) act += g; }
      const hrs = act / 3600000;
      return { w, scans: by[w].scans, pcs: by[w].pcs, min: Math.round(act / 60000), perHour: act >= 60000 ? Math.round(by[w].pcs / hrs) : null };
    });
    const sl = S.syncLog, okS = sl.filter(x => x.ok), ms = okS.map(x => x.ms);
    return { workers, errs, vUse, vFix, vFail, manual, sync: { n: sl.length, fail: sl.length - okS.length, avg: ms.length ? Math.round(ms.reduce((a, b) => a + b, 0) / ms.length) : null, max: ms.length ? Math.max.apply(null, ms) : null } };
  }
  const ERRNAME = { notInOrder: '注文にない商品', over: '数量オーバー', wrongStore: '店違い・担当外', badqr: 'QR違い', nomaster: 'マスター未登録' };
  V.output = {
    title: '出荷明細・作業記録',
    html() {
      if (!S.orders) return '<p class="hint">受注が未取込です。</p>';
      const m = metrics();
      const wr = m.workers.map(x => '<tr><td>' + esc(x.w) + '</td><td class="n">' + x.pcs + '枚</td><td class="n">' + x.min + '分</td><td class="n">' + (x.perHour == null ? '—' : x.perHour + '/時') + '</td></tr>').join('');
      const er = Object.keys(ERRNAME).map(k => '<tr><td>' + ERRNAME[k] + '</td><td class="n">' + (m.errs[k] || 0) + '回</td></tr>').join('') + '<tr><td>QRを使わず手入力で記録した回数</td><td class="n">' + m.manual + '回</td></tr>';
      return '<button class="btn green" data-act="print">出荷明細を印刷 / PDF保存</button>' +
        '<button class="btn gray small" data-act="paper">紙の予備を印刷(総量ピッキング表+手書き用出荷明細)</button>' +
        '<button class="btn gray small" data-act="backup">バックアップを保存・共有</button>' +
        '<button class="btn gray" data-act="csvShip">出荷明細(店別の最終数量)をCSVで保存</button>' +
        '<button class="btn gray" data-act="csvLog">作業記録(誰が・いつ・何を・何枚・どの店へ)をCSVで保存</button>' +
        '<div class="card"><b>試験運用の測定</b>' +
        '<table><tr><th>作業者</th><th class="n">枚数</th><th class="n">作業時間</th><th class="n">1人あたり</th></tr>' + (wr || '<tr><td colspan="4">まだ記録がありません</td></tr>') + '</table>' +
        '<p class="hint">作業時間は、操作の間隔が5分以内の時間だけを数えています。</p>' +
        '<table><tr><th>読み取りミス</th><th class="n">回数</th></tr>' + er + '</table>' +
        '<table style="margin-top:10px"><tr><th>音声</th><th class="n"></th></tr><tr><td>音声で入れた回数</td><td class="n">' + m.vUse + '</td></tr><tr><td>聞き間違い(数字を直した)</td><td class="n">' + m.vFix + '</td></tr><tr><td>音声が使えなかった</td><td class="n">' + m.vFail + '</td></tr></table>' +
        '<table style="margin-top:10px"><tr><th>通信</th><th class="n"></th></tr><tr><td>送信した回数 / 失敗</td><td class="n">' + m.sync.n + ' / ' + m.sync.fail + '</td></tr><tr><td>平均の遅れ</td><td class="n">' + (m.sync.avg == null ? '—' : m.sync.avg + 'ミリ秒') + '</td></tr><tr><td>最大の遅れ</td><td class="n">' + (m.sync.max == null ? '—' : m.sync.max + 'ミリ秒') + '</td></tr></table></div>' +
        '<button class="btn gray small" data-act="csvMetrics">測定記録をCSVで保存</button>';
    }
  };

  // ---- 設定 ----
  V.settings = {
    title: '設定',
    html() {
      const st = S.orders ? '<div class="card"><b>担当の店(チェックした店だけ仕分け対象。全部外すと全店)</b>' +
        S.orders.stores.map(s => '<label style="display:flex;gap:12px;align-items:center;margin:10px 0;font-size:22px"><input type="checkbox" style="width:30px;height:30px" data-chg="assign" data-s="' + esc(s) + '"' + (S.assigned.indexOf(s) >= 0 ? ' checked' : '') + '>' + esc(s) + '</label>').join('') + '</div>' : '<p class="hint">受注を取り込むと、担当の店を選べます。</p>';
      return st +
        '<div class="card"><b>記録の送信先(任意)</b><p class="hint">Googleスプレッドシート用のURL。空欄なら、スマホの中とCSVだけに記録します。電波がある時だけ自動で送ります。</p>' +
        '<input type="url" data-chg="syncurl" value="' + esc(S.syncUrl) + '" placeholder="https://script.google.com/macros/s/…/exec">' +
        '<button class="btn gray small" data-act="syncNow">いま送信する</button></div>' +
        '<div class="card"><b>仕分けをやり直す</b><p class="hint">進捗を0に戻します(受注はそのまま。これまでの作業記録は残ります)。</p><button class="btn red small" data-act="resetProg">進捗をリセット</button></div>';
    }
  };

  // ================= 印刷(出荷明細) =================
  function printShip() {
    const dm = doneMap();
    $('#printArea').innerHTML = S.orders.stores.map(s => {
      const st = storeStat(s);
      return '<div class="pg"><h2>出荷明細 — ' + esc(s) + '</h2><div>' + fmt(Date.now()) + ' / 作業者: ' + esc(S.worker) + ' / 受注: ' + esc(S.orders.file) + '</div>' +
        '<table style="margin-top:8px"><tr><th>品番</th><th>サイズ</th><th>色</th><th class="n">注文</th><th class="n">出荷</th><th class="n">差</th></tr>' +
        Object.keys(S.orders.need[s]).sort().map(k => { const p = parseKey(k), n = S.orders.need[s][k], d = dm[s + '|' + k] || 0; return '<tr><td>' + p.hin + '</td><td>' + p.size + '</td><td>' + esc(colorName(p.col)) + '</td><td class="n">' + n + '</td><td class="n">' + d + '</td><td class="n">' + (d - n) + '</td></tr>'; }).join('') +
        '<tr><th colspan="3">合計</th><th class="n">' + st.need + '</th><th class="n">' + st.done + '</th><th class="n">' + (st.done - st.need) + '</th></tr></table></div>';
    }).join('');
    setTimeout(() => window.print(), 50);
  }

  // ================= 記録の送信(電波がある時だけ) =================
  let syncing = false;
  async function trySync(manual) {
    if (!S.syncUrl) { if (manual) toast('送信先URLが未設定です(設定画面)'); return; }
    if (syncing) return;
    const pend = S.log.filter(e => !e.synced).slice(0, 200);
    if (!pend.length) { if (manual) toast('送信するものはありません'); return; }
    if (!navigator.onLine) { if (manual) toast('電波がありません。電波がある所で自動送信します'); return; }
    syncing = true; const t0 = performance.now(); let ok = false;
    try {
      const ac = new AbortController(); const to = setTimeout(() => ac.abort(), 15000);
      const r = await fetch(S.syncUrl, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ entries: pend }), signal: ac.signal });
      clearTimeout(to); const j = await r.json();
      if (j && j.ok) { ok = true; pend.forEach(e => { e.synced = true; DB.putLog(e); }); }
    } catch (e) { ok = false; }
    S.syncLog.push({ ts: Date.now(), ms: Math.round(performance.now() - t0), n: pend.length, ok }); if (S.syncLog.length > 500) S.syncLog.shift();
    persist('syncLog'); syncing = false; updNet();
    if (manual) toast(ok ? pend.length + '件 送信しました' : '送信できませんでした(あとで自動で再送します)');
  }

  // ================= もしもの備え =================
  // --- QRが読めない時: 商品一覧から選ぶ ---
  function manualList() {
    const dm = doneMap(), rem = {};
    visibleStores().forEach(s => Object.keys(S.orders.need[s]).forEach(k => { const r = S.orders.need[s][k] - (dm[s + '|' + k] || 0); if (r > 0) rem[k] = (rem[k] || 0) + r; }));
    const keys = Object.keys(rem).sort();
    return '<details class="card"><summary style="font-size:20px;font-weight:700">QRが読めない時はここ(商品を選ぶ)</summary>' +
      (keys.length ? keys.map(k => '<button class="btn gray small" data-act="manualKey" data-k="' + k + '">' + esc(nameOf(k)) + '(残り ' + rem[k] + '枚)</button>').join('') : '<p class="hint">残りの商品はありません</p>') + '</details>';
  }

  // --- バックアップ / 引き継ぎ ---
  const okCount = () => S.log.filter(e => e.kind === 'ok').length;
  function updBk() {
    const b = $('#bkbar'); if (!b) return;
    const lb = S.lastBackup, unsaved = okCount() - (lb ? lb.count : 0);
    const synced = S.syncUrl && !S.log.some(e => !e.synced && e.kind === 'ok');
    const stale = !lb || Date.now() - lb.ts > 30 * 60000;
    if (S.view === 'check' || !(unsaved > 0 && stale) || synced) { b.hidden = true; return; }
    b.innerHTML = '⚠ 記録' + unsaved + '件が、まだ別の場所に保存されていません<br><button class="chip" data-act="backup">いま保存・共有</button>';
    b.hidden = false;
  }
  function backupObj() { return { app: 'yamada-picking', v: 1, ts: Date.now(), worker: S.worker, master: S.master, colors: S.colors, colmap: S.colmap, orders: S.orders, assigned: S.assigned, picked: S.picked, log: S.log }; }
  async function saveBackup() {
    const name = 'yamada_backup_' + fmt(Date.now()).replace(/[-: ]/g, '').slice(0, 12) + '_' + (S.worker || 'noname') + '.json';
    const text = JSON.stringify(backupObj()); let done = false;
    try {
      const f = new File([text], name, { type: 'application/json' });
      if (navigator.canShare && navigator.canShare({ files: [f] })) { await navigator.share({ files: [f], title: 'ヤマダ出荷バックアップ' }); done = true; }
    } catch (e) { if (e && e.name === 'AbortError') { toast('保存をやめました'); return; } }
    if (!done) {
      const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([text], { type: 'application/json' }));
      a.download = name; document.body.appendChild(a); a.click(); a.remove();
    }
    S.lastBackup = { ts: Date.now(), count: okCount() }; persist('lastBackup'); updBk();
    toast('バックアップを保存しました(' + name + ')', 5000);
  }
  async function restoreBackup(o) {
    if (!o || o.app !== 'yamada-picking' || !Array.isArray(o.log)) throw new Error('ヤマダ出荷のバックアップではありません');
    const same = !!(S.orders && o.orders && S.orders.batch === o.orders.batch);
    if (!same && S.orders && S.log.some(x => x.batch === S.orders.batch && x.kind === 'ok' && !x.undone)) {
      if (!confirm('今の仕分けとは別の受注のバックアップです。今の進捗は引き継ぎ内容に置き換わります(作業記録は残ります)。続けますか?')) return null;
    }
    const byId = {}; S.log.forEach(e => { byId[e.id] = e; });
    let added = 0;
    o.log.forEach(e => {
      const c = byId[e.id];
      if (!c) { e.synced = false; S.log.push(e); byId[e.id] = e; DB.putLog(e); added++; }
      else if (e.undone && !c.undone) { c.undone = true; c.synced = false; DB.putLog(c); }
    });
    if (o.orders) { S.orders = o.orders; S.picked = Object.assign({}, o.picked || {}, same ? S.picked : {}); if (!same) S.assigned = o.assigned || []; }
    if (o.master && o.master.length) S.master = o.master;
    if (o.colors && Object.keys(o.colors).length) S.colors = o.colors;
    if (o.colmap) S.colmap = o.colmap;
    for (const k of ['master', 'colors', 'colmap', 'orders', 'assigned', 'picked']) await persist(k);
    return { added, same };
  }

  // --- 紙の予備(手書きで出荷を続けるための印刷物) ---
  function printPaper() {
    const tot = {};
    S.orders.stores.forEach(s => Object.keys(S.orders.need[s]).forEach(k => { tot[k] = (tot[k] || 0) + S.orders.need[s][k]; }));
    let h = '<div class="pg"><h2>総量ピッキング表(紙の予備)</h2><div>' + fmt(Date.now()) + ' / 受注: ' + esc(S.orders.file) + '</div><table style="margin-top:8px"><tr><th>済</th><th>品番</th><th>サイズ</th><th>色</th><th class="n">合計</th></tr>' +
      Object.keys(tot).sort().map(k => { const p = parseKey(k); return '<tr><td style="width:14mm">☐</td><td>' + p.hin + '</td><td>' + p.size + '</td><td>' + esc(colorName(p.col)) + '</td><td class="n">' + tot[k] + '</td></tr>'; }).join('') + '</table></div>';
    h += S.orders.stores.map(s => '<div class="pg"><h2>出荷明細(手書き用) — ' + esc(s) + '</h2><div>' + fmt(Date.now()) + ' / 受注: ' + esc(S.orders.file) + '</div>' +
      '<table style="margin-top:8px"><tr><th>品番</th><th>サイズ</th><th>色</th><th class="n">注文</th><th class="n" style="width:26mm">実績(手書き)</th><th>メモ</th></tr>' +
      Object.keys(S.orders.need[s]).sort().map(k => { const p = parseKey(k); return '<tr style="height:11mm"><td>' + p.hin + '</td><td>' + p.size + '</td><td>' + esc(colorName(p.col)) + '</td><td class="n">' + S.orders.need[s][k] + '</td><td></td><td></td></tr>'; }).join('') +
      '</table><p style="margin-top:14mm">確認者 ______________ 　日時 ______________</p></div>').join('');
    $('#printArea').innerHTML = h; setTimeout(() => window.print(), 50);
  }

  // --- 朝の点検 ---
  async function runChecks() {
    const R = [], add = (name, st, msg) => R.push({ name, st, msg });
    try {
      await DB.set('_t', Date.now()); const v = await DB.get('_t');
      add('スマホ内の保存', DB.db && v ? 'ok' : 'warn', DB.db && v ? '保存できます' : '一時保存のみです(ブラウザを閉じると消えます。プライベートモードを避けてください)');
    } catch (e) { add('スマホ内の保存', 'ng', '保存できません'); }
    try {
      if (navigator.storage && navigator.storage.estimate) { const e = await navigator.storage.estimate(), free = (e.quota - e.usage) / 1048576; add('空き容量', free > 200 ? 'ok' : free > 50 ? 'warn' : 'ng', '約' + Math.round(free) + 'MB'); }
      if (navigator.storage && navigator.storage.persisted) { const p = await navigator.storage.persisted(); add('データの自動削除対策', p ? 'ok' : 'warn', p ? '保護されています' : '容量不足の時に消される可能性があります。「ホーム画面に追加」すると安全です'); }
    } catch (e) {}
    add('受注データ', S.orders ? (new Date(S.orders.loadedAt).toDateString() === new Date().toDateString() ? 'ok' : 'warn') : 'ng',
      S.orders ? S.orders.stores.length + '店 / ' + S.orders.total + '枚(' + fmt(S.orders.loadedAt).slice(5, 16) + ' 取込)' + (new Date(S.orders.loadedAt).toDateString() === new Date().toDateString() ? '' : ' ← 今日の受注ですか?') : '取り込まれていません');
    add('商品マスター', S.master.length ? 'ok' : 'warn', S.master.length ? S.master.length + '品' : '未読込(未登録商品の警告が出ません)');
    let sw = false; try { sw = !!(navigator.serviceWorker && navigator.serviceWorker.controller); } catch (e) {}
    add('電波なしで開く準備', sw ? 'ok' : 'warn', sw ? '準備できています' : 'まだです。電波のある所でこのページを開き直してください');
    try {
      const devs = navigator.mediaDevices ? (await navigator.mediaDevices.enumerateDevices()).filter(d => d.kind === 'videoinput') : [];
      if (!devs.length) add('カメラ', 'ng', 'カメラが見つかりません(手入力で続けられます)');
      else { const st = await navigator.mediaDevices.getUserMedia({ video: true }); st.getTracks().forEach(t => t.stop()); add('カメラ', 'ok', '使えます'); }
    } catch (e) { add('カメラ', 'ng', 'カメラが許可されていません(ブラウザの設定でカメラを許可してください)'); }
    try { if (navigator.getBattery) { const b = await navigator.getBattery(), lv = Math.round(b.level * 100); add('電池', b.charging || lv >= 50 ? 'ok' : lv >= 20 ? 'warn' : 'ng', lv + '%' + (b.charging ? '(充電中)' : '')); } } catch (e) {}
    if (navigator.onLine) {
      try { const r = await fetch(location.pathname, { method: 'HEAD', cache: 'no-store' }), d = Date.parse(r.headers.get('Date')); if (d) { const df = Math.round((Date.now() - d) / 1000); add('時計', Math.abs(df) < 120 ? 'ok' : 'ng', Math.abs(df) < 120 ? 'ずれ' + df + '秒' : 'スマホの時計が' + df + '秒ずれています(記録の時刻が狂います)'); } } catch (e) {}
    }
    add('通信', 'ok', navigator.onLine ? '電波あり' : '電波なし(照合はこのまま続けられます)');
    add('画面の自動消灯防止', 'wakeLock' in navigator ? 'ok' : 'warn', 'wakeLock' in navigator ? '仕分け中は画面を消しません' : 'この端末では画面が暗くなることがあります');
    const lb = S.lastBackup;
    add('バックアップ', lb || !okCount() ? 'ok' : 'warn', lb ? fmt(lb.ts).slice(5, 16) + ' に保存' : (okCount() ? 'まだ保存していません' : '記録なし'));
    const ng = R.filter(x => x.st === 'ng').length, wn = R.filter(x => x.st === 'warn').length;
    addLog({ kind: 'check', errType: (ng ? 'NG' + ng : wn ? 'WARN' + wn : 'OK') + ' ' + R.filter(x => x.st !== 'ok').map(x => x.name).join(',') });
    const box = $('#chk'); if (!box) return;
    box.innerHTML = '<div class="card"><div class="huge ' + (ng ? 'ng' : 'ok') + '">' + (ng ? '❌ 直してください' : wn ? '⚠️ 注意あり' : '✅ 始められます') + '</div>' +
      (ng ? '<p class="hint">直せない時は、予備のスマホか紙の予備を使ってください。</p>' : '') + '</div>' +
      R.map(x => '<div class="card"><div class="row"><div class="big" style="flex:none;width:2em">' + (x.st === 'ok' ? '✅' : x.st === 'warn' ? '⚠️' : '❌') + '</div><div><b>' + esc(x.name) + '</b><div class="status">' + esc(x.msg) + '</div></div></div></div>').join('');
  }
  V.check = {
    title: '朝の点検',
    html() { return '<div id="chk"><p class="hint">点検しています…(カメラの許可を聞かれたら「許可」)</p></div><button class="btn gray small" data-act="beep">🔊 音を鳴らして確かめる</button><button class="btn small" data-act="go" data-v="check">もう一度点検する</button>'; },
    init() { updBk(); runChecks(); }
  };

  // --- 画面を消さない(仕分け中) / 想定外エラーの自動記録 ---
  let wl = null;
  async function keepAwake(on) { try { if (on && navigator.wakeLock) wl = await navigator.wakeLock.request('screen'); else if (wl) { wl.release(); wl = null; } } catch (e) {} }
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible' && S.view === 'sort') keepAwake(true); });
  let sysErr = 0;
  const logSys = m => { if (sysErr++ < 20) addLog({ kind: 'sys_err', errType: String(m).slice(0, 160) }); };
  window.addEventListener('error', e => logSys(e.message || 'error'));
  window.addEventListener('unhandledrejection', e => logSys((e.reason && e.reason.message) || 'promise'));

  // ================= 操作の受け口 =================
  const A = {
    go: el => { const v = el.dataset.v; if ((v === 'pick' || v === 'sort' || v === 'progress' || v === 'output') && !needOrders()) return; if ((v === 'pick' || v === 'sort') && !needWorker()) return; go(v); },
    sample: async () => {
      S.master = DEF_MASTER.map(r => ({ hin: r[0], size: r[1], col: r[2], name: '' })); S.colors = Object.assign({}, DEF_COLORS);
      S.colmap = { store: 1, hin: 2, size: 3, col: 4, qty: 5, header: 'auto' };
      await persist('master'); await persist('colors'); await persist('colmap');
      await setOrders(buildOrders(parseCSV(SAMPLE_ORDERS), 'サンプル(テスト9点・3店)'));
      toast('サンプルを読み込みました'); go('import');
    },
    ptoggle: el => { const k = el.dataset.k; S.picked[k] = !S.picked[k]; persist('picked'); go('pick'); },
    pexp: el => { const s = el.dataset.s; S.exp[s] = !S.exp[s]; go('progress'); },
    manual: () => { const i = $('#manual'); if (i && i.value.trim()) { Snd.init(); onCode(i.value.trim().toUpperCase(), 'typed'); i.value = ''; } },
    pickstore: el => { P.store = el.dataset.s; renderPanel(); },
    setq: el => { P.qty = +el.dataset.n; renderPanel(); },
    incq: () => { P.qty = Math.min(999, P.qty + 1); renderPanel(); },
    decq: () => { P.qty = Math.max(1, P.qty - 1); renderPanel(); },
    allq: () => { const c = P.list.find(x => x.store === P.store); P.qty = c ? c.remain : 1; renderPanel(); },
    voice: () => startVoice(),
    confirm: () => confirmQty(false),
    cancel: () => { P = null; renderPanel(); flash('ok', 'やめました', 1200); },
    undo: () => {
      const e = S.log.filter(x => x.kind === 'ok' && !x.undone && x.batch === S.orders.batch).slice(-1)[0];
      if (!e) return; e.undone = true; e.synced = false; DB.putLog(e); Snd.beep(500, 0.15); flash('ok', '取り消しました: ' + esc(nameOf(e.key)) + ' ' + e.qty + '枚(' + esc(e.store) + ')', 3000); renderPanel(); trySync();
    },
    print: () => printShip(),
    paper: () => { if (needOrders()) printPaper(); },
    backup: () => saveBackup(),
    manualKey: el => { onCode(el.dataset.k, 'manual'); },
    beep: () => { Snd.init(); Snd.ok(); setTimeout(Snd.err, 700); toast('「ピッ」のあと「ブー」が鳴れば正常です(鳴らない時は音量・マナーモードを確認)', 5000); },
    csvShip: () => download('出荷明細_' + fmt(Date.now()).replace(/[: ]/g, '').slice(0, 12) + '.csv', toCSV(shipRows())),
    csvLog: () => {
      const rows = [['日時', '作業者', '結果', '店名', '品番', 'サイズ', '色コード', '数量', '入力方法', '読み取り方法(scan/manual/typed)', '音声の聞き取り', '音声を直した', 'エラー種別', '取消', 'ID']];
      S.log.slice().sort((a, b) => a.ts - b.ts).forEach(e => rows.push([fmt(e.ts), e.worker, e.kind, e.store || '', e.hin || '', e.size || '', e.col || '', e.qty || '', e.input || '', e.via || '', e.voiceHeard || '', e.voiceFixed ? '○' : '', e.errType || '', e.undone ? '○' : '', e.id]));
      download('作業記録_' + fmt(Date.now()).replace(/[: ]/g, '').slice(0, 12) + '.csv', toCSV(rows));
    },
    csvMetrics: () => {
      const m = metrics(), rows = [['項目', '値', '補足']];
      m.workers.forEach(x => { rows.push(['作業者 ' + x.w + ' 枚数', x.pcs, '']); rows.push(['作業者 ' + x.w + ' 作業分', x.min, '操作間隔5分以内']); rows.push(['作業者 ' + x.w + ' 枚/時', x.perHour == null ? '' : x.perHour, '']); });
      Object.keys(ERRNAME).forEach(k => rows.push(['読み取りミス ' + ERRNAME[k], m.errs[k] || 0, '']));
      rows.push(['手入力で記録した回数', m.manual, 'QRを使わなかった回数']); rows.push(['音声 使用', m.vUse, '']); rows.push(['音声 聞き間違い', m.vFix, '数字を直した回数']); rows.push(['音声 失敗', m.vFail, '']);
      rows.push(['通信 送信回数', m.sync.n, '']); rows.push(['通信 失敗', m.sync.fail, '']); rows.push(['通信 平均ms', m.sync.avg == null ? '' : m.sync.avg, '']); rows.push(['通信 最大ms', m.sync.max == null ? '' : m.sync.max, '']);
      download('測定記録_' + fmt(Date.now()).replace(/[: ]/g, '').slice(0, 12) + '.csv', toCSV(rows));
    },
    syncNow: () => trySync(true),
    resetProg: async () => {
      if (!confirm('仕分けの進捗を0に戻します。作業記録は残ります。よろしいですか?')) return;
      S.orders.batch = uid(); await persist('orders'); toast('進捗をリセットしました'); go('home');
    }
  };
  document.addEventListener('click', e => { const el = e.target.closest('[data-act]'); if (el && A[el.dataset.act]) { Snd.init(); A[el.dataset.act](el); } });
  document.addEventListener('keydown', e => { if (e.key === 'Enter' && e.target.id === 'manual') A.manual(); });
  $('#btnHome').addEventListener('click', () => go('home'));

  document.addEventListener('change', async e => {
    const t = e.target;
    if (t.dataset.chg === 'name') { S.worker = t.value.trim(); persist('worker'); }
    else if (t.dataset.chg === 'selstore') { S.selStore = t.value; }
    else if (t.dataset.chg === 'single') { S.single = t.checked; }
    else if (t.dataset.chg === 'syncurl') { S.syncUrl = t.value.trim(); persist('syncUrl'); updNet(); }
    else if (t.dataset.chg === 'assign') { const s = t.dataset.s, i = S.assigned.indexOf(s); if (t.checked && i < 0) S.assigned.push(s); if (!t.checked && i >= 0) S.assigned.splice(i, 1); persist('assigned'); }
    else if (t.dataset.chg === 'colmap') { const k = t.dataset.k; S.colmap[k] = k === 'header' ? t.value : Math.max(1, parseInt(t.value, 10) || 1); persist('colmap'); }
    else if (t.dataset.file === 'backup') {
      const f = t.files[0]; if (!f) return;
      try {
        const r = await restoreBackup(JSON.parse(await f.text()));
        if (r) { toast('引き継ぎました(追加' + r.added + '件' + (r.same ? '・同じ受注に合流' : '・受注も引き継ぎ') + ')', 6000); go('home'); }
      } catch (err) { toast('引き継げませんでした: ' + err.message, 6000); }
    }
    else if (t.dataset.file) {
      const f = t.files[0]; if (!f) return;
      try {
        const rows = await loadRows(f);
        if (t.dataset.file === 'master') {
          const body = rows[0] && /品番/.test(String(rows[0][0])) ? rows.slice(1) : rows;
          S.master = body.map(r => ({ hin: half(r[0]), size: half(r[1]).toUpperCase(), col: half(r[2]).toUpperCase(), name: String(r[3] || '').trim() })).filter(m => m.hin && m.size && m.col);
          await persist('master'); toast('商品マスター ' + S.master.length + '品を読み込みました');
        } else if (t.dataset.file === 'colors') {
          const body = rows[0] && /色コード/.test(String(rows[0][0])) ? rows.slice(1) : rows;
          S.colors = {}; body.forEach(r => { const k = half(r[0]).toUpperCase(); if (k) S.colors[k] = String(r[1] || '').trim(); });
          await persist('colors'); toast('色コード表 ' + Object.keys(S.colors).length + '色を読み込みました');
        } else {
          const hasProg = S.orders && S.log.some(x => x.batch === S.orders.batch && x.kind === 'ok' && !x.undone);
          if (hasProg && !confirm('今の仕分けの進捗が0に戻ります(作業記録は残ります)。新しい受注を取り込みますか?')) return;
          await setOrders(buildOrders(rows, f.name)); toast('受注を取り込みました');
        }
      } catch (err) { toast('読み込めませんでした: ' + err.message, 6000); }
      go('import');
    }
  });

  window.addEventListener('online', () => { updNet(); trySync(); });
  window.addEventListener('offline', updNet);
  setInterval(() => { updNet(); trySync(); }, 30000);

  // ================= 起動 =================
  async function boot() {
    await DB.open();
    for (const k of ['master', 'colors', 'colmap', 'orders', 'worker', 'assigned', 'picked', 'syncUrl', 'syncLog', 'lastBackup']) { const v = await DB.get(k); if (v != null) S[k] = v; }
    if (!Object.keys(S.colors).length) S.colors = Object.assign({}, DEF_COLORS);
    S.log = await DB.allLog();
    try { if (navigator.storage && navigator.storage.persist) navigator.storage.persist(); } catch (e) {}   // 容量不足でもブラウザに消されにくくする
    go('home');
    if ('serviceWorker' in navigator && /^https?:$/.test(location.protocol)) navigator.serviceWorker.register('sw.js').catch(() => {});
  }
  window.__app = { restoreBackup, backupObj, runChecks, printPaper, S, onCode, A, go, candidates, doneMap, buildOrders, parseCSV, parseJaNum, getP: () => P, setQty: n => { P.qty = n; }, confirmQty };
  boot();
})();
