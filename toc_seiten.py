# -*- coding: utf-8 -*-
"""Seitenzahlen für das Inhaltsverzeichnis aus dem gerenderten PDF lesen.

Word rechnet die PAGEREF-Felder beim Öffnen neu; im PDF stünden sonst die
alten Zahlen. Das Skript sucht jeden Verzeichniseintrag im PDF und schreibt
{Eintrag: Seite} als JSON — Eingabe für finalize_zytig.py --seiten.

Aufruf:  python3 toc_seiten.py DOKUMENT.docx GERENDERT.pdf ZIEL.json
"""

import json
import re
import sys
import zipfile

import fitz
from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def toc_titel(docx_pfad):
    """Verzeichniseinträge (ohne Seitenzahl) in Dokumentreihenfolge."""
    root = etree.fromstring(
        zipfile.ZipFile(docx_pfad).read("word/document.xml"))
    titel = []
    for par in root.iter(W + "p"):
        stil = par.find(W + "pPr/" + W + "pStyle")
        if stil is None or not stil.get(W + "val").startswith(("Verzeichnis",
                                                              "TOC")):
            continue
        texte = [t.text or "" for t in par.iter(W + "t")]
        if len(texte) < 2:
            continue
        titel.append("".join(texte[:-1]).strip())
    return titel


def normalisieren(text):
    return re.sub(r"\s+", " ", text).strip().lower()


def seiten_ermitteln(pdf_pfad, titel, ab_seite=3):
    dok = fitz.open(pdf_pfad)
    seiten = [normalisieren(s.get_text()) for s in dok]
    ergebnis = {}
    letzte = ab_seite
    for t in titel:
        gesucht = normalisieren(t)
        gefunden = None
        for nr in range(letzte, len(seiten)):
            if gesucht in seiten[nr]:
                gefunden = nr + 1          # PDF-Seite 1-basiert
                break
        if gefunden is None:               # rückwärts nachsehen
            for nr in range(ab_seite, letzte):
                if gesucht in seiten[nr]:
                    gefunden = nr + 1
                    break
        if gefunden:
            ergebnis[t] = gefunden
            letzte = gefunden - 1
        else:
            print("  nicht gefunden:", t)
    return ergebnis, len(seiten)


def main(docx_pfad, pdf_pfad, ziel):
    titel = toc_titel(docx_pfad)
    seiten, gesamt = seiten_ermitteln(pdf_pfad, titel)
    json.dump(seiten, open(ziel, "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    print("Seiten gesamt: %d  (durch 4 teilbar: %s, fehlen %d)"
          % (gesamt, gesamt % 4 == 0, (4 - gesamt % 4) % 4))
    for t in titel:
        print("  %-55s %s" % (t[:55], seiten.get(t, "?")))


if __name__ == "__main__":
    sys.exit(main(sys.argv[1], sys.argv[2], sys.argv[3]))
