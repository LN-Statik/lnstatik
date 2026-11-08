# LNStatik

Aplikácia na faktúry, odpracované hodiny, výdavky a akcie kapely – https://app.lnstatik.sk

Funguje samostatne, bez akejkoľvek služby tretej strany okrem vlastného Microsoft 365 a GitHubu:

| Čo | Kde |
|---|---|
| Aplikácia (tento repozitár) | GitHub Pages, doména `app.lnstatik.sk` (CNAME vo Webglobe → `ln-statik.github.io`) |
| Dáta | OneDrive LN-Statika: `Dokumenty/1-- LNStatik/4_LNStatik/2_Financie/LNStatik_data/lnstatik_data.json` |
| Denné zálohy dát | `.../LNStatik_data/zalohy/lnstatik_data_RRRR-MM-DD.json` (appka ich robí sama) |
| PDF faktúr | `.../LNStatik_zaloha/Faktury_PDF/<rok>/` (uloží sa pri každom stiahnutí / odoslaní) |
| Prílohy akcií | `.../LNStatik_data/Prilohy/` |
| Prihlásenie | Microsoft Entra ID, aplikácia „LNStatik“ (Client ID a Tenant ID v `config.js`) |

V repozitári nie sú žiadne osobné údaje – všetko je v OneDrive.

## Súbory
- `index.html` – celá aplikácia (HTML, CSS a JavaScript v jednom súbore)
- `config.js` – prepojenie s Microsoft 365
- `vendor/` – knižnice (MSAL, html2pdf, JSZip) a písma, aby appka nezávisela od cudzích serverov
- `sw.js`, `manifest.webmanifest`, `icons/` – inštalácia na plochu / do telefónu
- `src/` – zdroj úložiska v OneDrive a service workera

## Úpravy
Stačí upraviť `index.html` (priamo na GitHube cez ikonku ceruzky) a uložiť (Commit). Do pár minút je zmena na app.lnstatik.sk.

## Formát dát
`lnstatik_data.json` = `{ "items": { "<id>": {...} }, "u": { "<id>": čas_zmeny_ms }, "del": {...} }`.
Druh záznamu je v poli `kind`: `invoice`, `client`, `time`, `event`, `expense`, `settings`.

## Microsoft 365 – povolenia (Entra → App registrations → LNStatik → API permissions)
`User.Read`, `Files.ReadWrite`, `Mail.ReadWrite` (Delegated), Redirect URI typu SPA: `https://app.lnstatik.sk/`.
