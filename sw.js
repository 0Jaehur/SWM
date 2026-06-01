// SWMT Service Worker – push notifications + PWA
const CACHE = "swmt-v8";

self.addEventListener("install", () => { /* wait for SKIP_WAITING message before taking over */ });
self.addEventListener("activate", e => e.waitUntil(clients.claim()));

// Track an unread count so the home-screen app icon shows a number
let swBadgeCount = 0;
function swSetBadge(n) {
  swBadgeCount = Math.max(0, n | 0);
  if (self.navigator && "setAppBadge" in self.navigator) {
    (swBadgeCount > 0 ? self.navigator.setAppBadge(swBadgeCount)
                      : self.navigator.clearAppBadge()).catch(() => {});
  }
}

// Allow the page to trigger immediate takeover and sync the badge count
self.addEventListener("message", e => {
  if (!e.data) return;
  if (e.data.type === "SKIP_WAITING") self.skipWaiting();
  if (e.data.type === "SET_BADGE") swSetBadge(e.data.count);   // page is open and knows the real count
  if (e.data.type === "CLEAR_BADGE") swSetBadge(0);
});

self.addEventListener("push", e => {
  const d = e.data ? e.data.json() : {};
  // Bump the home-screen icon badge. Use the count from the payload if present,
  // otherwise increment so a closed app still shows "1", "2", …
  const next = (typeof d.badgeCount === "number") ? d.badgeCount : swBadgeCount + 1;
  e.waitUntil(
    Promise.all([
      self.registration.showNotification(d.title || "SWMT – New Post", {
        body: d.body || "New activity in Shoe Wedge Mafia Tour",
        icon: d.icon || "icon-192.png",
        badge: "badge-72.png",
        tag: "swmt-post",
        renotify: true,
        vibrate: [200, 100, 200],
        data: d
      }),
      Promise.resolve(swSetBadge(next))
    ])
  );
});

self.addEventListener("notificationclick", e => {
  e.notification.close();
  swSetBadge(0);
  e.waitUntil(
    clients.matchAll({ type: "window", includeUncontrolled: true }).then(cs => {
      for (const c of cs) {
        if ("focus" in c) return c.focus();
      }
      if (clients.openWindow) return clients.openWindow("./");
    })
  );
});
