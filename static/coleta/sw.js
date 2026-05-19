const CACHE_NAME = "agres-coleta-v16";
const ASSETS = [
  "./",
  "./index.html",
  "./sw.js",
  "./manifest.webmanifest",
  "./logo_agres.png"
];
const APP_SHELL = new URL("./index.html", self.location.href).href;

async function cacheAsset(cache, asset) {
  const url = new URL(asset, self.location.href);
  const request = new Request(url.href, { cache: "reload" });
  const response = await fetch(request);
  if (response.ok) {
    const shellCopy = response.clone();
    await cache.put(request, response);
    if (asset === "./" || asset === "./index.html") {
      await cache.put(APP_SHELL, shellCopy);
    }
  }
}

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => Promise.allSettled(ASSETS.map((asset) => cacheAsset(cache, asset))))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((key) =>
            ["agres-pages-offline-", "agres-offline-", "agres-coleta-"].some((prefix) => key.startsWith(prefix)) &&
            key !== CACHE_NAME
          )
          .map((key) => caches.delete(key))
      )
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;
  const request = event.request;
  const isDocument = request.mode === "navigate" || request.destination === "document";

  if (isDocument) {
    event.respondWith(
      fetch(request).then((response) => {
        const copy = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(request, copy));
        return response;
      }).catch(async () => {
        const cache = await caches.open(CACHE_NAME);
        return (
          await cache.match(request, { ignoreSearch: true })
        ) || (
          await cache.match(APP_SHELL, { ignoreSearch: true })
        ) || (
          await cache.match("./index.html", { ignoreSearch: true })
        );
      })
    );
    return;
  }

  event.respondWith(
    caches.match(request, { ignoreSearch: true }).then((cached) => cached || fetch(request).then((response) => {
      const copy = response.clone();
      caches.open(CACHE_NAME).then((cache) => cache.put(request, copy));
      return response;
    }).catch(() => caches.match("./index.html", { ignoreSearch: true })))
  );
});
