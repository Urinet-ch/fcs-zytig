# FCS Zytig

Vereinszeitung des FC Schattdorf als PDF — Layout im Design-System der
FC-Schattdorf-Website (`~/fc-schattdorf/DESIGN.md`): Vereinsfarben (`#e63124` / `#d9261c`),
Inter als einzige Schrift, scharfe Kanten bei interaktivem Chrome, leicht gerundete
Inhaltsflächen.

## Struktur

- `content.mjs` — **einzige Stelle für Inhalte** (Texte, Bildpfade, Termine, Resultate,
  Sponsoren). Aktuell mit `[PLATZHALTER: ...]`-Texten befüllt.
- `templates.mjs` — HTML-Struktur der 4 Zytig-Seiten.
- `style.css` — Layout/Typografie nach FCS-Design-System.
- `assets/img/` — Logo, später weitere Bilder.
- `build.mjs` — rendert `content.mjs` + `templates.mjs` zu HTML und druckt es via
  Playwright/Chromium zu `output/fcs-zytig.pdf`.

## PDF bauen (Docker)

```bash
docker compose run --rm build
```

Ergebnis liegt danach in `output/fcs-zytig.pdf`.

## Lokal ohne Docker

```bash
npm install
npx playwright install --with-deps chromium
npm run build
```

## Inhalte einsetzen

Echte Texte/Bilder/Resultate/Termine in `content.mjs` eintragen (Platzhalter `[...]`
ersetzen), Bilder nach `assets/img/` legen und Pfad im jeweiligen `bild`-Feld eintragen.
Layout/CSS muss dafür nicht angefasst werden.
