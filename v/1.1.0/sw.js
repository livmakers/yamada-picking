// 一度開いたページを、電波なしでも開けるようにする(看板の控え)。
// 版ごとに別の保存箱を使い、他の版の箱は消さない。新しい版は、アプリを閉じて開き直すまで有効にならない
// (作業中に勝手に入れ替わらない)。VER は deploy.ps1 が版番号に書き換える。
const VER = '1.1.0';
const SCOPE = new URL(self.registration.scope).pathname;
const CACHE = 'ymd|' + SCOPE + '|' + VER;
const FILES = ['./', 'index.html', 'style.css', 'app.js', 'manifest.webmanifest', 'lib/jsQR.js', 'lib/xlsx.mini.min.js'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES))); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => (k.indexOf('ymd|' + SCOPE + '|') === 0 && k !== CACHE) || (SCOPE.indexOf('/v/') < 0 && /^yamada-v\d+$/.test(k))).map(k => caches.delete(k)))));
});
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;
  if (/versions\.json$/.test(u.pathname)) return;   // 版の一覧は常に最新を取りに行く
  e.respondWith(caches.open(CACHE).then(c => c.match(e.request, { ignoreSearch: true }).then(hit => hit ||
    fetch(e.request).then(r => { if (r.ok) c.put(e.request, r.clone()); return r; }))));
});
