// SWMT Service Worker – push notifications + PWA
const CACHE = "swmt-v1";

self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", e => e.waitUntil(clients.claim()));

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
