// Service Worker do app de Inspeção de Campo
// Responsável por deixar o app carregando mesmo sem internet (cache do "shell" do app).
// Os DADOS preenchidos ficam separados, no IndexedDB (ver index.html) — isso aqui só
// garante que a PÁGINA em si abra offline.

const CACHE_NAME = 'inspecao-app-v4'; // aumente o número (v2, v3...) sempre que atualizar o app
const ARQUIVOS_PARA_CACHE = [
  './',
  './index.html',
  './dashboard.html',
  './manifest.json',
  './logo.png',
  './chart.umd.min.js'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ARQUIVOS_PARA_CACHE))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((nomes) =>
      Promise.all(
        nomes
          .filter((nome) => nome !== CACHE_NAME)
          .map((nome) => caches.delete(nome))
      )
    )
  );
  self.clients.claim();
});

// Estratégia: tenta a rede primeiro (pra pegar versão nova quando online);
// se falhar (sem internet), usa o que está em cache.
self.addEventListener('fetch', (event) => {
  // Nunca interceptar chamadas para o Google Apps Script (dados) — só o "shell" do app.
  if (event.request.method !== 'GET') return;

  event.respondWith(
    fetch(event.request)
      .then((resposta) => {
        const respostaClone = resposta.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, respostaClone));
        return resposta;
      })
      .catch(() => caches.match(event.request).then((r) => r || caches.match('./index.html')))
  );
});
