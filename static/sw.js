// 설비 QR 이력관리 - 오프라인 조회용 서비스 워커
// 전략: GET 요청만 캐시 대상. POST(고장등록/가동 시작·정지·종료/로그인 등)는 절대 캐시하지 않고 항상 네트워크로 보냄.
//       네트워크 우선(Network-First) - 성공하면 최신 데이터로 캐시 갱신, 실패(오프라인)하면 캐시된 마지막 화면을 보여줌.
const CACHE_NAME = 'equipment-qr-cache-v2';
const APP_SHELL = [
  '/static/manifest.json',
  '/static/icon-192.png',
  '/static/icon-512.png',
  '/static/img/logo.jpg'
];

self.addEventListener('install', (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return Promise.all(
        APP_SHELL.map((url) => cache.add(url).catch(() => {}))
      );
    })
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((names) =>
      Promise.all(names.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n)))
    ).then(() => self.clients.claim())
  );
});

const OFFLINE_FALLBACK = `<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>오프라인</title>
<style>body{font-family:'Malgun Gothic',sans-serif;background:#eef2f7;display:flex;align-items:center;
justify-content:center;height:100vh;margin:0;text-align:center;color:#444}
.box{padding:32px}</style></head><body><div class="box">
<div style="font-size:2.4rem;">📡</div>
<h3>오프라인 상태입니다</h3>
<p>네트워크 연결을 확인해주세요.<br>이전에 열어본 화면은 캐시에서 볼 수 있습니다.</p>
</div></body></html>`;

self.addEventListener('fetch', (event) => {
  const req = event.request;

  // POST 등 쓰기 요청은 절대 캐시하지 않고 그대로 네트워크로 전달 (고장이력/가동률 데이터 정합성 보호)
  if (req.method !== 'GET') {
    return;
  }

  // 다른 도메인(Bootstrap/Google Fonts 등 CDN) 요청은 가로채지 않고 브라우저 기본 동작에 맡긴다.
  // 서비스워커가 대신 처리하다가 배포 중 순간적인 네트워크 실패라도 겹치면 CSS/JS 전체가
  // 깨진 채로 캐시 없이 실패해버리는 위험이 있어, 우리가 통제하는 같은 오리진 리소스만 다룬다.
  if (!req.url.startsWith(self.location.origin)) {
    return;
  }

  event.respondWith(
    fetch(req)
      .then((res) => {
        if (res && res.ok) {
          const clone = res.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(req, clone));
        }
        return res;
      })
      .catch(async () => {
        const cached = await caches.match(req);
        if (cached) return cached;
        if (req.mode === 'navigate') {
          return new Response(OFFLINE_FALLBACK, { headers: { 'Content-Type': 'text/html; charset=utf-8' } });
        }
        return Response.error();
      })
  );
});
