# -*- coding: utf-8 -*-
"""Korrekturmodus-Änderungen (Tracked Changes) in einer .docx annehmen.

Word-Semantik beim «Alle Änderungen annehmen»:
  w:ins            → Inhalt bleibt, Marker verschwindet
  w:del            → Inhalt verschwindet
  w:trPr/w:del     → ganze Tabellenzeile verschwindet
  w:pPr/w:rPr/w:del→ Absatzmarke gelöscht → Absatz wird mit dem folgenden vereint
  *Change          → alte Formatierung verwerfen, neue behalten

Aufruf:  python3 docx_revisions.py QUELLE.docx ZIEL.docx
"""

import shutil
import sys
import zipfile

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
CHANGE_TAGS = ("rPrChange", "pPrChange", "tblPrChange", "trPrChange",
               "tcPrChange", "sectPrChange", "tblGridChange",
               "numberingChange")


def _unwrap(el):
    """Kind-Elemente an die Stelle des Elements setzen, Element entfernen."""
    parent = el.getparent()
    idx = parent.index(el)
    for child in reversed(list(el)):
        parent.insert(idx, child)
    parent.remove(el)


def _merge_with_next(par):
    """Absatz mit gelöschter Absatzmarke in den Folgeabsatz einhängen."""
    parent = par.getparent()
    idx = parent.index(par)
    nxt = None
    for sib in parent[idx + 1:]:
        if sib.tag == W + "p":
            nxt = sib
            break
        if sib.tag == W + "tbl":  # kein Folgeabsatz auf gleicher Ebene
            break
    inhalt = [c for c in par if c.tag != W + "pPr"]
    if nxt is None:
        # Nichts zum Vereinen: Absatzmarke einfach behalten.
        return False
    if inhalt:
        ppr = nxt.find(W + "pPr")
        pos = 1 if ppr is not None else 0
        for child in reversed(inhalt):
            nxt.insert(pos, child)
    parent.remove(par)
    return True


def accept_revisions(root):
    """Alle Revisionen im Dokumentbaum annehmen. → Statistik-Dict."""
    stat = dict.fromkeys(
        ("ins", "del", "zeilen", "absatzmarken", "format"), 0)

    # 1. Gelöschte Tabellenzeilen entfernen
    for trpr in list(root.iter(W + "trPr")):
        if trpr.find(W + "del") is not None:
            row = trpr.getparent()
            row.getparent().remove(row)
            stat["zeilen"] += 1

    # 2. Formatierungs-Revisionen: alte Formatierung verwerfen
    for tag in CHANGE_TAGS:
        for el in list(root.iter(W + tag)):
            el.getparent().remove(el)
            stat["format"] += 1

    # 3. Absatzmarken-Revisionen
    marken = []
    for ppr in root.iter(W + "pPr"):
        rpr = ppr.find(W + "rPr")
        if rpr is None:
            continue
        if rpr.find(W + "ins") is not None:      # eingefügte Marke → bleibt
            rpr.remove(rpr.find(W + "ins"))
        d = rpr.find(W + "del")
        if d is not None:
            rpr.remove(d)
            marken.append(ppr.getparent())
    for par in marken:
        if par.getparent() is not None and _merge_with_next(par):
            stat["absatzmarken"] += 1

    # 4. Gelöschte Inhalte entfernen, eingefügte behalten
    for el in list(root.iter(W + "del")):
        if el.getparent() is not None:
            el.getparent().remove(el)
            stat["del"] += 1
    for el in list(root.iter(W + "moveFrom")):
        if el.getparent() is not None:
            el.getparent().remove(el)
    for el in list(root.iter(W + "ins")):
        if el.getparent() is not None:
            _unwrap(el)
            stat["ins"] += 1
    for el in list(root.iter(W + "moveTo")):
        if el.getparent() is not None:
            _unwrap(el)

    # 5. Leere Tabellenzellen brauchen mindestens einen Absatz
    for tc in root.iter(W + "tc"):
        if tc.find(W + "p") is None:
            tc.append(etree.SubElement(tc, W + "p"))
    return stat


def main(quelle, ziel):
    shutil.copyfile(quelle, ziel)
    with zipfile.ZipFile(quelle) as z:
        namen = z.namelist()
        inhalte = {n: z.read(n) for n in namen}
    root = etree.fromstring(inhalte["word/document.xml"])
    stat = accept_revisions(root)
    inhalte["word/document.xml"] = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(ziel, "w", zipfile.ZIP_DEFLATED) as out:
        for n in namen:
            out.writestr(n, inhalte[n])
    print("Angenommen:", stat)


if __name__ == "__main__":
    sys.exit(main(sys.argv[1], sys.argv[2]))
