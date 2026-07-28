# FCS Zytig

Vereinszeitung des FC Schattdorf als PDF — Layout im Design-System der
FC-Schattdorf-Website (`~/fc-schattdorf/DESIGN.md`): Vereinsfarben (`#e63124` / `#d9261c`),
Inter als einzige Schrift, scharfe Kanten bei interaktivem Chrome, leicht gerundete
Inhaltsflächen.

## Ausgabe Sommer 2026 — Stand nach der Korrekturrunde

Die Ausgabe entsteht heute als **Word-Dokument** (nicht mehr über die HTML-Strecke
weiter unten):

1. `build_zytig_docx.py` — erzeugt den Entwurf aus den Inhalten in
   `Zytig Sommer 26/`. Achtung: der abgelieferte Entwurf wurde danach in Word
   von Hand nachbearbeitet, das Skript bildet diesen Stand nicht mehr 1:1 ab.
2. `docx_revisions.py` — nimmt Korrekturmodus-Änderungen einer .docx an
   (`python3 docx_revisions.py QUELLE.docx ZIEL.docx`).
3. `finalize_zytig.py` — Schlussbearbeitung: übernimmt die Korrekturen von
   Ralph Bomatter aus `Zytig Sommer 26/FCS-Zyttig Sommer 2026 ENTWURF - rbo.docx`,
   ersetzt Titelseiten-Logo, Brückli-Logo und Teamfoto FCS 3, ergänzt die
   Grümpi-Bilder, räumt die Redaktionshinweise weg und verkleinert die Bilder
   auf Druckauflösung (Datei bleibt so per Mail versendbar).
4. `toc_seiten.py` — liest die Seitenzahlen des Inhaltsverzeichnisses aus dem
   gerenderten PDF, damit sie auch ohne Word-Feldaktualisierung stimmen.

```bash
python3 finalize_zytig.py --fueller 1            # Word-Datei bauen
python3 toc_seiten.py DOK.docx DOK.pdf seiten.json
python3 finalize_zytig.py --fueller 1 --seiten seiten.json
```

Das PDF entsteht aus dem fertigen .docx über Word («Sichern unter» → PDF).

## Struktur (ältere HTML-Strecke)

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
