/* 宅建合格（配信版）SW — build.pyが3d6a084c2aを本体ハッシュで置換する。
   index.htmlは【ネットワーク優先】= 更新が即反映。オフライン時のみキャッシュで起動。 */
const CACHE = "takken-app-3d6a084c2a";
const ASSETS = ["./", "./index.html", "./manifest.webmanifest", "./icon-512.png", "./apple-touch-icon.png"];
self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener("fetch", e => {
  const req = e.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  const isDoc = req.mode === "navigate" || url.pathname.endsWith("/index.html") || url.pathname.endsWith("/study/");
  if (isDoc) {
    // ネットワーク優先（最新を取得しキャッシュ更新）。圏外時のみキャッシュ
    e.respondWith(fetch(req).then(res => {
      const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy));
      return res;
    }).catch(() => caches.match(req).then(h => h || caches.match("./index.html"))));
  } else if (url.origin === location.origin) {
    e.respondWith(caches.match(req).then(hit => hit || fetch(req).then(res => {
      if (res && res.status === 200) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
      return res;
    })));
  }
});
