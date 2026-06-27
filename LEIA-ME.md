# Músicas e Acordes — PWA (instalar no telemóvel)

Tal como a colecção de vinis: alojas os ficheiros uma vez e instalas no telemóvel
com "Adicionar ao ecrã principal". Funciona offline e guarda as tuas músicas, listas
e favoritas no próprio telemóvel.

## Ficheiros (mantém todos juntos, na mesma pasta)
- `index.html` — a aplicação
- `manifest.json` — identidade da app (nome, ícone LF, cores)
- `sw.js` — service worker (offline)
- `icon-192.png`, `icon-512.png` — ícones

## Porque precisa de alojamento
Para instalar e ter offline, os ficheiros têm de ser servidos por **HTTPS**.
Abrir o `index.html` diretamente do telemóvel (file://) mostra a app, mas **não**
deixa instalar nem funcionar offline. O alojamento faz-se **uma única vez**.

## Opção A — Netlify Drop (a mais rápida)
1. Vai a https://app.netlify.com/drop
2. Arrasta a pasta inteira para a página.
3. Recebes um endereço https (ex.: `https://nome-qualquer.netlify.app`).

## Opção B — GitHub Pages
1. Cria um repositório e envia estes ficheiros.
2. Settings → Pages → Branch `main` / `/root` → Save.
3. Fica em `https://utilizador.github.io/repositorio/`.

## Opção C — Cloudflare Pages
1. https://pages.cloudflare.com → Create → Direct Upload.
2. Carrega a pasta. Recebes um endereço https.

## Instalar no Android
1. Abre o endereço https no **Chrome**.
2. Menu (⋮) → **Instalar aplicação** / **Adicionar ao ecrã principal**.
3. Passa a abrir em ecrã inteiro, com o ícone LF próprio.

## Instalar no iPhone (iOS)
1. Abre o endereço no **Safari**.
2. Partilhar → **Adicionar ao ecrã principal**.

## Notas honestas
- **O backup verdadeiro é o "Exportar" dentro da app**, não o armazenamento do
  telemóvel. Se desinstalares a app ou limpares os dados do browser, as músicas/listas
  que criaste vão com eles. Exporta o JSON de vez em quando e guarda o ficheiro.
- **Atualizar a app:** quando eu te der um `index.html` novo, sobe os ficheiros outra
  vez e muda `CACHE = "musicas-acordes-v1"` para `"...-v2"` no `sw.js`, para forçar a
  atualização nos telemóveis.
- Esta app não usa internet nenhuma — depois de instalada funciona sempre offline.
