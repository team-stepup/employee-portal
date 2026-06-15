// employee-portal Service Worker
// 戦略: 静的アセット (HTML/JS/CSS/画像) は network-first → 失敗時に cache。
//      API (Azure Functions) は network-only (キャッシュしない)。
// 更新: 「🔄 更新」ボタンを押すと skipWaiting() で即時切替。

const CACHE_NAME = 'portal-v28';
const STATIC_ASSETS = ['/', '/index.html', '/manifest.json', '/logo.png', '/emblem.png', '/icon-192.png', '/icon-512.png', '/pdf-lib.min.js', '/zaishoku_jp.png', '/seal_stepup.png'];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(STATIC_ASSETS).catch(() => {}))
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => Promise.all(
      keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k))
    )).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);
  // API は SW を介さない (network-only)
  if (url.hostname.includes('azurewebsites.net')) return;
  if (event.request.method !== 'GET') return;

  // 静的アセットは network-first
  event.respondWith(
    fetch(event.request).then((resp) => {
      if (resp && resp.status === 200 && resp.type === 'basic') {
        const respClone = resp.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, respClone)).catch(() => {});
      }
      return resp;
    }).catch(() => caches.match(event.request))
  );
});

// 更新メッセージ: クライアントから SKIP_WAITING を受け取ったら即時切替
self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});

// ====== Web Push ======
self.addEventListener('push', (event) => {
  let data = {};
  try { data = event.data ? event.data.json() : {}; } catch (e) { data = { body: event.data ? event.data.text() : '' }; }
  const title = data.title || '従業員ポータル';
  const options = {
    body: data.body || '',
    icon: '/icon-192.png',
    badge: '/icon-192.png',
    data: { url: data.url || '/' },
    tag: data.tag || 'portal-notice',
    renotify: true,
  };
  const tasks = [self.registration.showNotification(title, options)];
  // ホーム画面アイコンのバッジを設定 (対応端末・iOS16.4+のPWA)
  if (self.navigator && 'setAppBadge' in self.navigator) {
    const n = typeof data.badge === 'number' ? data.badge : 1;
    tasks.push(self.navigator.setAppBadge(n).catch(() => {}));
  }
  event.waitUntil(Promise.all(tasks));
});

self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  const target = (event.notification.data && event.notification.data.url) || '/';
  event.waitUntil(
    self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then((list) => {
      for (const c of list) {
        if ('focus' in c) { c.navigate(target); return c.focus(); }
      }
      if (self.clients.openWindow) return self.clients.openWindow(target);
    })
  );
});
