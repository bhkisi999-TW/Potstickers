// 一煎鍾情 Service Worker：基本離線支援
// 修改快取內容時，請更換 VERSION，舊快取會在啟用時自動清除。
const VERSION = 'yijiao-v5';
const CORE = [
  './',
  './index.html',
  './manifest.webmanifest',
  './assets/story-night.jpg',
  './assets/reserve-gallery.jpg',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/icon-maskable-512.png',
  './icons/apple-touch-icon.png',
  './icons/favicon-32.png',
  './icons/favicon-48.png',
  './assets/menu/gyoza-pan.jpg',
  './assets/menu/gyoza-boiled.jpg',
  './assets/menu/gift-box.jpg',
  './assets/menu/soup-hot-sour.jpg',
  './assets/menu/cucumber.jpg',
  './assets/menu/tofu.jpg',
  './assets/menu/gyoza-leek.jpg',
  './assets/menu/gyoza-sichuan.jpg',
  './assets/menu/shrimp-dumpling.jpg',
  './assets/menu/soymilk.jpg'
];
// 外部資源（套件、圖片）：先用快取、背景更新
const EXTERNAL_HOSTS = ['fonts.googleapis.com', 'fonts.gstatic.com'];
const EXTERNAL_MAX = 60;
const PRECACHE_EXTERNAL = ['https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@300;400;500;700&display=swap'];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(VERSION).then((cache) =>
      Promise.allSettled([
        ...CORE.map((url) => cache.add(url)),
        // 網頁字型樣式（跨網域，用 no-cors 先存起來；字型檔本身會在第一次連網瀏覽時自動存入；失敗不影響安裝）
        ...PRECACHE_EXTERNAL.map((url) => fetch(new Request(url, { mode: 'no-cors' })).then((res) => cache.put(url, res)))
      ])
    ).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== VERSION).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

async function trim(cache) {
  const keys = await cache.keys();
  const ext = keys.filter((k) => new URL(k.url).origin !== self.location.origin);
  if (ext.length > EXTERNAL_MAX) {
    await Promise.all(ext.slice(0, ext.length - EXTERNAL_MAX).map((k) => cache.delete(k)));
  }
}

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // 頁面導覽：網路優先（確保拿到最新版），離線時用快取
  if (req.mode === 'navigate') {
    event.respondWith(
      fetch(req).then((res) => {
        const copy = res.clone();
        caches.open(VERSION).then((c) => c.put('./index.html', copy));
        return res;
      }).catch(() => caches.match('./index.html').then((r) => r || caches.match('./')))
    );
    return;
  }

  // 同網域資源：快取優先，沒有再連網並存起來
  if (url.origin === self.location.origin) {
    event.respondWith(
      caches.match(req).then((hit) => hit || fetch(req).then((res) => {
        if (res.ok) { const copy = res.clone(); caches.open(VERSION).then((c) => c.put(req, copy)); }
        return res;
      }))
    );
    return;
  }

  // 指定的外部資源：先給快取、背景更新
  if (EXTERNAL_HOSTS.includes(url.hostname)) {
    event.respondWith(
      caches.open(VERSION).then(async (cache) => {
        const hit = await cache.match(req);
        const net = fetch(req).then((res) => {
          if (res && (res.ok || res.type === 'opaque')) { cache.put(req, res.clone()).then(() => trim(cache)); }
          return res;
        }).catch(() => hit);
        return hit || net;
      })
    );
  }
});
