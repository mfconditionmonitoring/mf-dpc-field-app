// Offline cache for the MF-DPC-132 Field App
const CACHE = "mf-dpc-132-c066b83b38";
const FILES = ["./", "./index.html", "./manifest.json", "./icon-192.png", "./icon-512.png"];
self.addEventListener("install", (e) => { e.waitUntil(caches.open(CACHE).then((c) => c.addAll(FILES)).then(() => self.skipWaiting())); });
self.addEventListener("activate", (e) => { e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener("fetch", (e) => {
  // only the app's own files; uploads and the office link always go to the network
  if (e.request.method !== "GET" || new URL(e.request.url).origin !== self.location.origin || /\/exec$/.test(new URL(e.request.url).pathname)) return;
  // network first so updates arrive when there is signal; cache when offline
  e.respondWith(fetch(e.request).then((r) => { const cp = r.clone(); caches.open(CACHE).then((c) => c.put(e.request, cp)); return r; }).catch(() => caches.match(e.request, { ignoreSearch: true }).then((r) => r || caches.match("./index.html"))));
});
