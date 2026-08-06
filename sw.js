// Service Worker do app de Inspeção de Campo
// Responsável por deixar o app carregando mesmo sem internet (cache do "shell" do app).
// Os DADOS preenchidos ficam separados, no IndexedDB (ver index.html) — isso aqui só
// garante que a PÁGINA em si abra offline.

const CACHE_NAME = 'inspecao-app-v13'; // aumente o número sempre que atualizar o app
const ARQUIVOS_PARA_CACHE = [
  './',
  './index.html',
  './manifest.json',
  './logo.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(async (cache) => {
      // Guarda cada arquivo SEPARADAMENTE (em vez de cache.addAll, que é "tudo ou nada").
      // Se um arquivo falhar, os outros continuam sendo salvos normalmente — evita a
      // tela em branco que acontecia quando um único arquivo da lista dava erro.
      await Promise.all(
        ARQUIVOS_PARA_CACHE.map(async (url) => {
          try {
            const resposta = await fetch(url, { cache: 'no-store' });
            if (resposta.ok) {
              await cache.put(url, resposta);
            } else {
              console.warn('Não consegui guardar em cache (resposta não OK):', url);
            }
          } catch (erro) {
            console.warn('Não consegui guardar em cache (falhou o fetch):', url, erro);
          }
        })
      );
    })
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
  if (event.request.method !== 'GET') return;

  event.respondWith(
    fetch(event.request, { cache: 'no-store' })
      .then((resposta) => {
        const respostaClone = resposta.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, respostaClone));
        return resposta;
      })
      .catch(() => caches.match(event.request).then((r) => r || caches.match('./index.html')))
  );
});
