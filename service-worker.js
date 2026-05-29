// employee-portal Service Worker
// 戦略: 静的アセット (HTML/JS/CSS/画像) は network-first → 失敗時に cache。
//      API (Azure Functions) は network-only (キャッシュしない)。
// 更新: 「🔄 更新」ボタンを押すと skipWaiting() で即時切替。

const CACHE_NAME = 'portal-v11';
const STATIC_ASSETS = ['/', '/index.html', '/manifest.json', '/logo.png', '/emblem.png', '/icon-192.png', '/icon-512.png'];

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
