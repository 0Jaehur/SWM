// SWMT Service Worker – push notifications + PWA
const CACHE = "swmt-v7";

self.addEventListener("install", () => { /* wait for SKIP_WAITING message before taking over */ });
self.addEventListener("activate", e => e.waitUntil(clients.claim()));

// Allow the page to trigger immediate takeover
self.addEventListener("message", e => {
  if (e.data && e.data.type === "SKIP_WAITING") self.skipWaiting();
});

self.addEventListener("push", e => {
  const d = e.data ? e.data.json() : {};
  e.waitUntil(
    self.registration.showNotification(d.title || "SWMT – New Post", {
      body: d.body || "New activity in Shoe Wedge Mafia Tour",
      icon: d.icon || "icon-192.png",
      badge: "badge-72.png",
      tag: "swmt-post",
      renotify: true,
      vibrate: [200, 100, 200],
      data: d
    })
  );
});

self.addEventListener("notificationclick", e => {
  e.notification.close();
  e.waitUntil(
    clients.matchAll({ type: "window", includeUncontrolled: true }).then(cs => {
      for (const c of cs) {
        if ("focus" in c) return c.focus();
      }
      if (clients.openWindow) return clients.openWindow("./");
    })
  );
});
