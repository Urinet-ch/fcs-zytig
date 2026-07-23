const bildPlatz = (bild, text, klasse = '') =>
  bild
    ? `<div class="bild-platz ${klasse}"><img src="${bild}" alt=""></div>`
    : `<div class="bild-platz ${klasse}">${text}</div>`;

const masthead = (c) => `
  <header class="masthead">
    <div class="brand">
      <img class="logo" src="assets/img/fcs-logo.svg" alt="FC Schattdorf Logo">
      <div>
        <div class="titel">FCS Zytig</div>
        <div class="strapline">${c.ausgabe.strapline}</div>
      </div>
    </div>
    <div class="meta">Ausgabe ${c.ausgabe.nummer}<br>${c.ausgabe.datum}</div>
  </header>
  <div class="accent-line"></div>
`;

const footer = (c) => `
  <div class="footer-band">
    <div class="verein">${c.impressum.verein}</div>
    <div class="kontakt">${c.impressum.adresse} · ${c.impressum.kontakt}</div>
  </div>
`;

const seite1 = (c) => `
  <section class="page">
    ${masthead(c)}
    <article class="lead">
      <span class="kategorie">${c.leitartikel.kategorie}</span>
      <h1>${c.leitartikel.titel}</h1>
      <p class="dachzeile">${c.leitartikel.dachzeile}</p>
      ${bildPlatz(c.leitartikel.bild, c.leitartikel.bildtext)}
      <div class="fliesstext">
        ${c.leitartikel.absaetze.map((p) => `<p>${p}</p>`).join('\n')}
      </div>
    </article>
  </section>
`;

const seite2 = (c) => `
  <section class="page">
    ${masthead(c)}
    <div>
      <h2 class="section-titel">News</h2>
      <div class="news-grid">
        ${c.news
          .map(
            (n) => `
          <article class="news-karte">
            <div class="bild-platz">
              <div class="badges">
                <span class="badge-datum">${n.datum}</span>
                <span class="badge-kategorie">${n.kategorie}</span>
              </div>
              ${n.bild ? `<img src="${n.bild}" alt="">` : ''}
            </div>
            <div class="inhalt">
              <h3>${n.titel}</h3>
              <p>${n.text}</p>
            </div>
          </article>`
          )
          .join('\n')}
      </div>
    </div>
    <div class="resultate-band">
      <h2 class="section-titel">${c.resultate.titel}</h2>
      <div class="resultate-tabelle">
        ${c.resultate.zeilen
          .map(
            (r) => `
          <div class="resultate-zeile">
            <span>${r.team} – ${r.gegner}</span>
            <span class="resultat">${r.resultat}</span>
          </div>`
          )
          .join('\n')}
      </div>
    </div>
  </section>
`;

const seite3 = (c) => `
  <section class="page">
    ${masthead(c)}
    <div>
      <h2 class="section-titel">Agenda</h2>
      <div class="agenda-liste">
        ${c.agenda
          .map(
            (a) => `
          <div class="agenda-eintrag">
            <div class="datum">${a.datum}<br>${a.zeit}</div>
            <div>
              <div class="titel">${a.titel}</div>
              <div class="ort">${a.ort}</div>
            </div>
          </div>`
          )
          .join('\n')}
      </div>
    </div>
    <div class="junioren-box">
      <h2 class="section-titel">${c.junioren.titel}</h2>
      <p>${c.junioren.text}</p>
    </div>
  </section>
`;

const seite4 = (c) => `
  <section class="page">
    ${masthead(c)}
    <div>
      <h2 class="section-titel">Sponsoren</h2>
      <div class="sponsoren-grid">
        ${c.sponsoren.map((s) => `<div class="sponsor-kachel">${s}</div>`).join('\n')}
      </div>
    </div>
    ${footer(c)}
  </section>
`;

export function renderHTML(c) {
  return `<!doctype html>
<html lang="de-CH">
<head>
<meta charset="utf-8">
<title>FCS Zytig — Ausgabe ${c.ausgabe.nummer}</title>
<link rel="stylesheet" href="style.css">
</head>
<body>
${seite1(c)}
${seite2(c)}
${seite3(c)}
${seite4(c)}
</body>
</html>`;
}
