// Taller Gary (app): la página queda guardada en el teléfono y abre AL INSTANTE con la copia guardada.
// Por detrás se baja la versión publicada; si cambió (APP_BUILD distinto), se avisa a la página para recargar.
const CACHE = "taller-gary-app";
const build = t => ((t || "").match(/const APP_BUILD='([^']+)'/) || [])[1] || "";

self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", e => e.waitUntil(self.clients.claim()));

self.addEventListener("fetch", e => {
  const r = e.request;
  if (r.method !== "GET" || r.mode !== "navigate") return;           // datos de Google, imágenes, etc.: normal
  const url = new URL(r.url);
  if (url.origin !== location.origin) return;
  e.respondWith((async () => {
    const c = await caches.open(CACHE), clave = new URL("./", self.registration.scope).href;
    const guardada = await c.match(clave);
    const copia = guardada ? guardada.clone() : null;     // la original se entrega a la pantalla
    const red = (async () => {
      try {
        const res = await fetch(clave, { cache: "no-store" });
        if (!res.ok) return null;
        const texto = await res.clone().text();
        await c.put(clave, res.clone());
        const antes = copia ? build(await copia.text()) : "";
        const nuevo = build(texto);
        if (nuevo) await c.put(new URL("build.txt", self.registration.scope).href, new Response(nuevo));   // la página lo revisa al abrir
        if (guardada && nuevo && nuevo !== antes) {                       // hay versión nueva publicada
          for (const w of await self.clients.matchAll({ type: "window" })) w.postMessage({ nuevaVersion: nuevo });
        }
        return res;
      } catch (_) { return null; }
    })();
    if (guardada) { e.waitUntil(red); return guardada; }               // al instante; se actualiza por detrás
    return (await red) || fetch(r);
  })());
});
