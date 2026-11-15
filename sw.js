// Offline cache: the app works on the shop floor without coverage.
const CACHE = "status-field-v0.8.1";
const FILES = ["./", "./index.html", "./style.css", "./app.js", "./zip.js", "./manifest.webmanifest", "./icons/icon-192.png", "./icons/icon-512.png"];
self.addEventListener("install", e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES))); self.skipWaiting(); });
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))));
  self.clients.claim();
});
// Android: "Share" from WhatsApp / Files / Gmail -> Status Field (manifest share_target).
// The file is kept in a cache and the app opens it (references .json or report .zip).
self.addEventListener("fetch", e => {
  if (e.request.method !== "POST" || !new URL(e.request.url).pathname.endsWith("/share-target")) return;
  e.respondWith((async () => {
    try {
      const form = await e.request.formData();
      const f = form.getAll("file").find(x => x && typeof x === "object");
      if (f) {
        const c = await caches.open("status-field-shared");
        await c.put("./shared-file", new Response(f, { headers: { "X-File-Name": encodeURIComponent(f.name || "shared") } }));
      }
    } catch (err) { /* opens the app anyway */ }
    return Response.redirect("./index.html?shared=1", 303);
  })());
});

// Network first (new versions arrive at once), cache when there is no coverage
self.addEventListener("fetch", e => {
  if (e.request.method !== "GET") return;
  e.respondWith(fetch(e.request).then(r => {
    const copy = r.clone();
    caches.open(CACHE).then(c => c.put(e.request, copy));
    return r;
  }).catch(() => caches.match(e.request).then(r => r || caches.match("./index.html"))));
});
