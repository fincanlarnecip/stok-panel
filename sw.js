// Basit servis çalışanı: uygulama kabuğunu önbelleğe alır, PWA kurulabilirliğini sağlar.
// Veri her zaman canlı (Apps Script) API'sinden çekilir, önbelleğe alınmaz.
const CACHE_ADI = "stok-panel-v1";
const ONBELLEK_DOSYALARI = ["./", "./index.html", "./manifest.json", "./icon-192.png", "./icon-512.png"];

self.addEventListener("install", (e) => {
  e.waitUntil(
    caches.open(CACHE_ADI).then((cache) => cache.addAll(ONBELLEK_DOSYALARI))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_ADI).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (e) => {
  // Sadece kendi dosyalarımızı önbellekten sun (Apps Script API isteklerine dokunma)
  if (e.request.method !== "GET" || !e.request.url.startsWith(self.location.origin)) return;
  e.respondWith(
    caches.match(e.request).then((cached) => cached || fetch(e.request))
  );
});
