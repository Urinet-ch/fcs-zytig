// Einzige Stelle für Inhalte. Layout (templates.mjs / style.css) bleibt unangetastet,
// wenn hier später die echten Texte/Bilder des FC Schattdorf eingesetzt werden.

export const content = {
  ausgabe: {
    nummer: '01',
    datum: '[PLATZHALTER: Ausgabe-Datum, z. B. Sommer 2026]',
    strapline: '[PLATZHALTER: Untertitel der Zytig]',
  },

  leitartikel: {
    kategorie: '[Rubrik]',
    titel: '[PLATZHALTER: Haupttitel der Ausgabe]',
    dachzeile: '[PLATZHALTER: kurzer Anriss / Dek-Satz zum Haupttitel]',
    bild: null, // Pfad zu assets/img/... sobald vorhanden
    bildtext: '[Bildlegende]',
    absaetze: [
      '[PLATZHALTER: erster Absatz des Leitartikels. Dieser Text wird später durch den ' +
        'echten Bericht ersetzt.]',
      '[PLATZHALTER: zweiter Absatz.]',
    ],
  },

  news: [
    {
      datum: '[TT.MM.]',
      kategorie: '[Rubrik]',
      titel: '[PLATZHALTER: News-Titel 1]',
      text: '[PLATZHALTER: kurzer News-Text.]',
      bild: null,
    },
    {
      datum: '[TT.MM.]',
      kategorie: '[Rubrik]',
      titel: '[PLATZHALTER: News-Titel 2]',
      text: '[PLATZHALTER: kurzer News-Text.]',
      bild: null,
    },
    {
      datum: '[TT.MM.]',
      kategorie: '[Rubrik]',
      titel: '[PLATZHALTER: News-Titel 3]',
      text: '[PLATZHALTER: kurzer News-Text.]',
      bild: null,
    },
  ],

  resultate: {
    titel: 'Resultate & Tabelle',
    zeilen: [
      { team: '[Team]', gegner: '[Gegner]', resultat: '[x:x]' },
      { team: '[Team]', gegner: '[Gegner]', resultat: '[x:x]' },
      { team: '[Team]', gegner: '[Gegner]', resultat: '[x:x]' },
    ],
  },

  agenda: [
    { datum: '[TT.MM.]', zeit: '[hh:mm]', titel: '[PLATZHALTER: Termin]', ort: '[Ort]' },
    { datum: '[TT.MM.]', zeit: '[hh:mm]', titel: '[PLATZHALTER: Termin]', ort: '[Ort]' },
    { datum: '[TT.MM.]', zeit: '[hh:mm]', titel: '[PLATZHALTER: Termin]', ort: '[Ort]' },
    { datum: '[TT.MM.]', zeit: '[hh:mm]', titel: '[PLATZHALTER: Termin]', ort: '[Ort]' },
  ],

  junioren: {
    titel: 'Junioren-Ecke',
    text: '[PLATZHALTER: Text zu einem Juniorenteam / Nachwuchs-Thema.]',
  },

  sponsoren: [
    '[Sponsor 1]', '[Sponsor 2]', '[Sponsor 3]', '[Sponsor 4]',
    '[Sponsor 5]', '[Sponsor 6]',
  ],

  impressum: {
    verein: 'FC Schattdorf',
    adresse: '[PLATZHALTER: Adresse]',
    kontakt: '[PLATZHALTER: E-Mail / Web]',
  },
};
