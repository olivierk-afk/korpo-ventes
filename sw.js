self.addEventListener('install', (event) => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(self.clients.claim());
});

// Réception d'une notification Web Push envoyée par Supabase quand un représentant est mentionné (@)
self.addEventListener('push', (event) => {
  let payload = {
    title: '💬 Nouvelle mention dans Korpo Top Earner',
    body: 'Quelqu’un vous a identifié dans le canal d’équipe.',
    msgId: '',
    url: './index.html'
  };
  try {
    if (event.data) {
      const parsed = event.data.json();
      payload = { ...payload, ...parsed };
    }
  } catch (e) {
    if (event.data) payload.body = event.data.text();
  }

  event.waitUntil(
    self.registration.showNotification(payload.title, {
      body: payload.body,
      tag: payload.msgId ? `korpo-mention-${payload.msgId}` : 'korpo-mention',
      renotify: true,
      data: { msgId: payload.msgId, url: payload.url || './index.html' }
    })
  );
});

// Ouverture directe de l'application au clic sur la notification
self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  const targetUrl = (event.notification.data && event.notification.data.url) || './index.html';
  event.waitUntil(
    self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then((clientList) => {
      for (const client of clientList) {
        if ('focus' in client) return client.focus();
      }
      if (self.clients.openWindow) return self.clients.openWindow(targetUrl);
    })
  );
});
