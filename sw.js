// Servis çalışanı: HTML sayfası için "önce ağ" (network-first) stratejisi kullanır,
// böylece uygulama her açılışta en güncel içeriği gösterir. Sadece internet
// bağlantısı yokken önbellekteki son bilinen sürüme düşer (offline yedek).
const CACHE_ADI = "stok-panel-v4";
const ONBELLEK_DOSYALARI = ["./manifest.json", "./icon-192.png", "./icon-512.png", "./icon-192-maskable.png", "./icon-512-maskable.png"];

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
  if (e.request.method !== "GET" || !e.request.url.startsWith(self.location.origin)) return;

  // Sayfa (HTML) istekleri: HER ZAMAN önce ağdan taze veri çekmeyi dene — cache:"no-store" ile
  // GitHub Pages'in kendi HTTP önbelleğini (Cache-Control max-age) de by-pass ediyoruz, yoksa
  // deploy'dan hemen sonra bile birkaç dakika eski sürüm gösterilebiliyordu.
  if (e.request.mode === "navigate" || e.request.url.endsWith("index.html") || e.request.url.endsWith("/stok-panel/")) {
    e.respondWith(
      fetch(e.request, { cache: "no-store" })
        .then((res) => {
          const kopya = res.clone();
          caches.open(CACHE_ADI).then((cache) => cache.put(e.request, kopya));
          return res;
        })
        .catch(() => caches.match(e.request))
    );
    return;
  }

  // Diğer statik dosyalar (ikon, manifest vb.): önce önbellek, yoksa ağdan çek.
  e.respondWith(
    caches.match(e.request).then((cached) => cached || fetch(e.request))
  );
});
