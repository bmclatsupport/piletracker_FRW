/* Service worker for the crew page.

   Caches the app shell so the page opens with no signal after the first visit.
   Pile data and map tiles are stored by the app itself in IndexedDB — not here.

   ---------------------------------------------------------------------------------------
   Why the page is fetched from the network first, and the icons are not.

   This used to be cache-first for everything, including the HTML document. That is fine
   until you publish a new build: the next visit is answered out of the old cache and only
   revalidated afterwards, so the first load after an upload shows the *previous* app and
   you have to reload to see what you just published. On a site with two builds on one
   origin that reads as the wrong app appearing at the right address, which is exactly how
   it looked — the crew page coming back as the coordinator app.

   So the document is network-first with a cache fallback: online you always get what is
   actually published, offline you get the last copy that worked. Icons and the manifest
   stay cache-first because they change once a year and are worth having instantly.

   The offline fallback is also deliberately scoped. It used to fall back to "./index.html"
   resolved against this worker's own location, which for the worker at the site root is
   the coordinator build — so a failed crew page request could be answered with the
   coordinator app. A worker now only ever falls back to a document it actually cached, and
   only for a request inside its own scope.
   --------------------------------------------------------------------------------------- */
const CACHE = "pilehq-crew-v34";
/* office.html is the coordinator page — same app, coordinator role, no project. Cached
   alongside the crew page so a coordinator on a phone keeps working when the signal drops,
   which on a jobsite is most of the time. hosting.cjs fails if any SHELL entry is missing,
   because addAll rejects the whole list if one file 404s. */
const SHELL = ["./", "./index.html", "./office.html", "./manifest.webmanifest",
  "./icon-180.png", "./icon-192.png", "./icon-512.png"];
/* everything this worker is allowed to answer for — its own folder and below */
const SCOPE = new URL("./", self.location).pathname;

self.addEventListener("install", e => {
  e.waitUntil((async () => {
    const c = await caches.open(CACHE);
    await Promise.allSettled(SHELL.map(u => c.add(new Request(u, { cache: "reload" })).catch(() => { })));
    self.skipWaiting();
  })());
});

self.addEventListener("activate", e => {
  e.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)));
    await self.clients.claim();
  })());
});

const isDoc = req => req.mode === "navigate"
  || (req.destination === "document")
  || (req.headers.get("accept") || "").indexOf("text/html") >= 0;

self.addEventListener("fetch", e => {
  const req = e.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  // map tiles and the DroneDeploy API are handled by the app; let them pass through
  if (/arcgisonline|openstreetmap|dronedeploy/.test(url.hostname)) return;
  if (url.origin !== location.origin) return;
  // never answer for a sibling build living in another folder on the same origin
  if (url.pathname.indexOf(SCOPE) !== 0) return;

  if (isDoc(req)) {
    e.respondWith((async () => {
      const c = await caches.open(CACHE);
      try {
        const r = await fetch(req, { cache: "no-store" });   // always the published build
        if (r && r.ok) c.put(req, r.clone());
        return r;
      } catch (err) {
        return (await c.match(req, { ignoreSearch: true }))
          || (await c.match("./index.html"))
          || new Response("Offline, and this page has not been saved on the device yet.",
            { status: 503, headers: { "Content-Type": "text/plain" } });
      }
    })());
    return;
  }

  e.respondWith((async () => {
    const c = await caches.open(CACHE);
    const hit = await c.match(req, { ignoreSearch: true });
    if (hit) {
      fetch(req).then(r => { if (r && r.ok) c.put(req, r.clone()); }).catch(() => { });
      return hit;                                  // instant, works offline
    }
    try {
      const r = await fetch(req);
      if (r && r.ok) c.put(req, r.clone());
      return r;
    } catch (err) {
      return new Response("", { status: 504 });
    }
  })());
});
