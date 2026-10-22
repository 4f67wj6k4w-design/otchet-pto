/* Отчёт ПТО: служебный файл приложения. Страница всегда берётся из сети (свежая версия), из памяти — только если сети нет. Данные из базы не кэшируются. */
const CACHE = "pto-shell-v1";
self.addEventListener("install", e => { self.skipWaiting(); });
self.addEventListener("activate", e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener("fetch", e => {
  const r = e.request, u = new URL(r.url);
  if (r.method !== "GET" || u.origin !== location.origin) return;          // база (script.google.com) и шрифты — мимо
  e.respondWith(fetch(r).then(res => { if (res.ok) { const c = res.clone(); caches.open(CACHE).then(ch => ch.put(r, c)); } return res; })
    .catch(() => caches.match(r, { ignoreSearch: true })));
});
