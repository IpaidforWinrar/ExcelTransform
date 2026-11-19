"""
trafoUpdate.py

Uebernimmt die Werte aus "Trafo Freeze-Liste.xlsx" in eine Kopie der
"empty Trafo Freeze-Liste.xlsx". Zugeordnet wird ueber den zusammengesetzten
Primary Key (Stellennummer + Personalnummer) im Reiter "SOLL Wertstromzuordnung".

Die Zielliste wird NICHT mit openpyxl neu gespeichert: Die Werte werden direkt in
die Sheet-XML der kopierten Datei geschrieben. Alle anderen Teile der Datei
(Datenverbindungen, Power Query, Slicer, Formatierung, Blattschutz, ...) bleiben
unveraendert. openpyxl wird nur zum Lesen der Keys und der Quellwerte benutzt.

Ordner (relativ zu diesem Skript):
    update Eingabe/   empty Trafo Freeze-Liste.xlsx
                      Trafo Freeze-Liste.xlsx
    update Ausgabe/   filled Trafo Freeze-Liste.xlsx   (wird erzeugt)
                      trafoUpdate_Protokoll.txt         (wird erzeugt)

Benoetigt: pip install openpyxl lxml
"""

import datetime
import os
import re
import shutil
import sys
import zipfile
from collections import defaultdict
from decimal import Decimal
from pathlib import Path, PurePosixPath

from lxml import etree
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.datetime import to_excel

# =============================================================================
# EINSTELLUNGEN - hier anpassen
# =============================================================================

# Ordner und Dateinamen
INPUT_DIR = "update Eingabe"
OUTPUT_DIR = "update Ausgabe"
EMPTY_FILE = "empty Trafo Freeze-Liste.xlsx"
SOURCE_FILE = "Trafo Freeze-Liste.xlsx"
FILLED_FILE = "filled Trafo Freeze-Liste.xlsx"
LOG_FILE = "trafoUpdate_Protokoll.txt"

# Reiter, in dem gematcht und kopiert wird
SHEET_NAME = "SOLL Wertstromzuordnung"

# Erste Datenzeile (alles darueber sind Kopfzeilen)
FIRST_DATA_ROW = 3

# Primary-Key-Spalten (zusammengesetzter Schluessel): Stellennummer, Personalnummer
KEY_COLUMNS = ["D", "F"]

# Spalten, die uebernommen werden. Einzelne Spalten und/oder Bereiche moeglich,
# z.B. ["I:BJ"]  oder  ["I:M", "P", "R:BJ"]
COPY_COLUMNS = ["I:BJ"]

# Spalten, die trotzdem NICHT uebernommen werden (gleiches Format), z.B. ["K", "AA:AC"]
EXCLUDE_COLUMNS = []

# Was tun, wenn ein Key mehrfach vorkommt (in Quelle oder Ziel)?
#   "skip"  = nicht uebernehmen, nur im Protokoll auflisten
#   "first" = jeweils die erste Zeile verwenden
DUPLICATE_POLICY = "skip"

# Gross-/Kleinschreibung beim Vergleichen der Keys ignorieren?
KEY_IGNORE_CASE = True

# Excel beim Oeffnen alle Formeln neu berechnen lassen? (empfohlen: Formeln, die
# auf die geaenderten Zellen verweisen, haben sonst noch alte Zwischenwerte)
RECALC_ON_OPEN = True

# =============================================================================

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_ILLEGAL_XML = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f]")


# ----------------------------------------------------------------------------
# Hilfsfunktionen
# ----------------------------------------------------------------------------

def parse_columns(specs):
    """['I:M', 'P'] -> [9, 10, 11, 12, 13, 16] (Spaltenindizes, sortiert, ohne Doppelte)"""
    result = set()
    for spec in specs:
        spec = spec.strip().upper()
        if ":" in spec:
            a, b = spec.split(":")
            lo, hi = column_index_from_string(a), column_index_from_string(b)
            result.update(range(min(lo, hi), max(lo, hi) + 1))
        else:
            result.add(column_index_from_string(spec))
    return sorted(result)


def norm(value):
    """Key-Wert vereinheitlichen: None -> '', 123.0 -> '123', Leerzeichen entfernen."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = str(value).strip()
    return text.lower() if KEY_IGNORE_CASE else text


def read_keys(ws, key_cols):
    """Liefert {key_tuple: [zeilennummern]} ab FIRST_DATA_ROW. Komplett leere Keys werden uebersprungen."""
    index = defaultdict(list)
    for row in range(FIRST_DATA_ROW, ws.max_row + 1):
        key = tuple(norm(ws.cell(row, c).value) for c in key_cols)
        if all(part == "" for part in key):
            continue
        index[key].append(row)
    return index


def fmt_key(key):
    return " | ".join(key) if len(key) > 1 else key[0]


def q(tag):
    return f"{{{NS}}}{tag}"


def col_of(ref):
    return column_index_from_string(re.match(r"[A-Z]+", ref).group())


# ----------------------------------------------------------------------------
# XML-Ebene
# ----------------------------------------------------------------------------

def find_sheet_part(zin, sheet_name):
    """Pfad der Sheet-XML im ZIP fuer den Reiternamen (ueber workbook.xml + rels)."""
    wb = etree.fromstring(zin.read("xl/workbook.xml"))
    rid = None
    for sheet in wb.iter(q("sheet")):
        if sheet.get("name") == sheet_name:
            rid = sheet.get(f"{{{NS_R}}}id")
    if rid is None:
        raise KeyError(f"Reiter '{sheet_name}' nicht in workbook.xml gefunden")
    rels = etree.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
    for rel in rels:
        if rel.get("Id") == rid:
            target = rel.get("Target")
            if target.startswith("/"):
                return target.lstrip("/")
            return str(PurePosixPath("xl") / target)
    raise KeyError(f"Relationship {rid} nicht gefunden")


def get_row(sheet_data, rows, r):
    """Zeilen-Element holen oder (an der richtigen Stelle) neu anlegen."""
    if r in rows:
        return rows[r]
    new = etree.Element(q("row"))
    new.set("r", str(r))
    nxt = next((el for n, el in sorted(rows.items()) if n > r), None)
    if nxt is not None:
        nxt.addprevious(new)
    else:
        sheet_data.append(new)
    rows[r] = new
    return new


def get_cell(row_el, r, c):
    """Zellen-Element holen oder (in Spaltenreihenfolge) neu anlegen. Stil bleibt erhalten."""
    nxt = None
    for el in row_el:
        if el.tag != q("c"):
            continue
        cc = col_of(el.get("r"))
        if cc == c:
            return el
        if cc > c:
            nxt = el
            break
    new = etree.Element(q("c"))
    new.set("r", f"{get_column_letter(c)}{r}")
    if nxt is not None:
        nxt.addprevious(new)
    else:
        row_el.append(new)
    return new


def write_value(c_el, src_cell, epoch):
    """Wert in <c> schreiben. Stil (s=...) bleibt, Formeln/alte Werte werden entfernt.
    Gibt zurueck: (hatte_Formel, war_Shared_Formula_Master)"""
    had_formula, shared_master = False, False
    for child in list(c_el):
        if child.tag == q("f"):
            had_formula = True
            shared_master = child.get("t") == "shared" and child.get("ref") is not None
        c_el.remove(child)
    if "t" in c_el.attrib:
        del c_el.attrib["t"]

    v = src_cell.value
    if v is None:
        return had_formula, shared_master

    if src_cell.data_type == "e":
        c_el.set("t", "e")
        etree.SubElement(c_el, q("v")).text = str(v)
    elif isinstance(v, bool):
        c_el.set("t", "b")
        etree.SubElement(c_el, q("v")).text = "1" if v else "0"
    elif isinstance(v, (int, float, Decimal)):
        etree.SubElement(c_el, q("v")).text = repr(v) if isinstance(v, float) else str(v)
    elif isinstance(v, (datetime.datetime, datetime.date, datetime.time, datetime.timedelta)):
        etree.SubElement(c_el, q("v")).text = repr(to_excel(v, epoch=epoch))
    else:
        c_el.set("t", "inlineStr")
        is_el = etree.SubElement(c_el, q("is"))
        t_el = etree.SubElement(is_el, q("t"))
        t_el.text = _ILLEGAL_XML.sub("", str(v))
        t_el.set(XML_SPACE, "preserve")
    return had_formula, shared_master


def set_full_calc(wb_xml):
    """<calcPr fullCalcOnLoad="1"/> setzen (an schema-konformer Stelle anlegen, falls fehlend)."""
    root = etree.fromstring(wb_xml)
    calc = root.find(q("calcPr"))
    if calc is None:
        calc = etree.Element(q("calcPr"))
        successors = {"oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes",
                      "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"}
        nxt = next((el for el in root if etree.QName(el).localname in successors), None)
        if nxt is not None:
            nxt.addprevious(calc)
        else:
            root.append(calc)
    calc.set("fullCalcOnLoad", "1")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def drop_calc_chain(zin):
    """calcChain.xml entfernen (Excel baut sie neu auf). Gibt geaenderte Teile zurueck."""
    changes = {}
    ct = etree.fromstring(zin.read("[Content_Types].xml"))
    for el in list(ct):
        if el.get("PartName") == "/xl/calcChain.xml":
            ct.remove(el)
    changes["[Content_Types].xml"] = etree.tostring(ct, xml_declaration=True, encoding="UTF-8", standalone=True)
    rels = etree.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
    for el in list(rels):
        if el.get("Type", "").endswith("/calcChain"):
            rels.remove(el)
    changes["xl/_rels/workbook.xml.rels"] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)
    return changes


def rewrite_zip(path, replacements, removed):
    """ZIP neu schreiben: alle Teile unveraendert, ausser replacements (Name->Bytes) und removed."""
    tmp = path.with_suffix(".tmp")
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w") as zout:
        for item in zin.infolist():
            if item.filename in removed:
                continue
            data = replacements.get(item.filename)
            if data is None:
                data = zin.read(item.filename)
            zout.writestr(item, data, compress_type=item.compress_type)
    os.replace(tmp, path)


# ----------------------------------------------------------------------------
# Hauptprogramm
# ----------------------------------------------------------------------------

def main():
    base = Path(__file__).resolve().parent
    in_dir, out_dir = base / INPUT_DIR, base / OUTPUT_DIR
    empty_path, source_path = in_dir / EMPTY_FILE, in_dir / SOURCE_FILE
    filled_path, log_path = out_dir / FILLED_FILE, out_dir / LOG_FILE

    for p in (empty_path, source_path):
        if not p.exists():
            sys.exit(f"Datei nicht gefunden: {p}")

    key_cols = parse_columns(KEY_COLUMNS)
    excluded = set(parse_columns(EXCLUDE_COLUMNS))
    copy_cols = [c for c in parse_columns(COPY_COLUMNS) if c not in excluded]
    if not copy_cols:
        sys.exit("Keine Spalten zum Kopieren konfiguriert (COPY_COLUMNS).")

    # 1) empty -> Ausgabe kopieren und umbenennen
    out_dir.mkdir(exist_ok=True)
    shutil.copyfile(empty_path, filled_path)
    print(f"Kopiert: {empty_path.name} -> {filled_path}")

    # 2) Nur LESEN mit openpyxl (wird nie gespeichert): Keys und Quellwerte
    wb_target = load_workbook(filled_path, data_only=True)
    wb_source = load_workbook(source_path, data_only=True)
    for name, wb in (("Ziel", wb_target), ("Quelle", wb_source)):
        if SHEET_NAME not in wb.sheetnames:
            sys.exit(f"Reiter '{SHEET_NAME}' fehlt in {name}-Datei. Vorhanden: {wb.sheetnames}")
    ws_t, ws_s = wb_target[SHEET_NAME], wb_source[SHEET_NAME]

    header_diffs = []
    for row in range(1, FIRST_DATA_ROW):
        for c in copy_cols:
            a, b = norm(ws_t.cell(row, c).value), norm(ws_s.cell(row, c).value)
            if a != b:
                header_diffs.append(f"{get_column_letter(c)}{row}: Ziel='{a}' / Quelle='{b}'")

    idx_t = read_keys(ws_t, key_cols)
    idx_s = read_keys(ws_s, key_cols)

    dup_t = {k: r for k, r in idx_t.items() if len(r) > 1}
    dup_s = {k: r for k, r in idx_s.items() if len(r) > 1}
    unmatched_t = {k: r for k, r in idx_t.items() if k not in idx_s}
    unmatched_s = {k: r for k, r in idx_s.items() if k not in idx_t}

    # 3) Zuordnung Zielzeile -> Quellzeile
    assignments = {}  # zielzeile -> quellzeile
    skipped_dups = 0
    for key, rows_t in idx_t.items():
        if key not in idx_s:
            continue
        rows_s = idx_s[key]
        if DUPLICATE_POLICY == "skip" and (len(rows_t) > 1 or len(rows_s) > 1):
            skipped_dups += len(rows_t)
            continue
        for tgt_row in rows_t:
            assignments[tgt_row] = rows_s[0]

    # 4) Werte direkt in die Sheet-XML der kopierten Datei schreiben
    cells_written = 0
    formulas_overwritten = 0
    shared_masters = []
    with zipfile.ZipFile(filled_path) as zin:
        sheet_part = find_sheet_part(zin, SHEET_NAME)
        root = etree.fromstring(zin.read(sheet_part))
        sheet_data = root.find(q("sheetData"))
        rows = {int(el.get("r")): el for el in sheet_data.findall(q("row"))}

        for tgt_row, src_row in sorted(assignments.items()):
            row_el = get_row(sheet_data, rows, tgt_row)
            row_el.attrib.pop("spans", None)  # nur ein Optimierungshinweis, wird von Excel neu gesetzt
            for c in copy_cols:
                if isinstance(ws_t.cell(tgt_row, c), MergedCell):
                    continue
                c_el = get_cell(row_el, tgt_row, c)
                had_f, was_master = write_value(c_el, ws_s.cell(src_row, c), wb_source.epoch)
                cells_written += 1
                if had_f:
                    formulas_overwritten += 1
                if was_master:
                    shared_masters.append(f"{get_column_letter(c)}{tgt_row}")

        replacements = {
            sheet_part: etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        }
        removed = set()
        if assignments and RECALC_ON_OPEN:
            replacements["xl/workbook.xml"] = set_full_calc(zin.read("xl/workbook.xml"))
        if formulas_overwritten and "xl/calcChain.xml" in zin.namelist():
            replacements.update(drop_calc_chain(zin))
            removed.add("xl/calcChain.xml")

    rewrite_zip(filled_path, replacements, removed)

    # 5) Protokoll
    def block(title, mapping, label):
        lines = [f"{title} ({len(mapping)})", "-" * 60]
        if not mapping:
            lines.append("  (keine)")
        for key, rows_ in sorted(mapping.items()):
            lines.append(f"  {fmt_key(key)}   [{label} Zeile {', '.join(map(str, rows_))}]")
        lines.append("")
        return lines

    key_names = " + ".join(get_column_letter(c) for c in key_cols)
    copy_names = ", ".join(get_column_letter(c) for c in copy_cols)
    log = [
        f"trafoUpdate Protokoll - {datetime.datetime.now():%d.%m.%Y %H:%M:%S}",
        f"Reiter: {SHEET_NAME}",
        f"Primary Key (Spalten): {key_names}  - Format im Protokoll: Stellennummer | Personalnummer",
        f"Kopierte Spalten: {copy_names}",
        f"Duplikat-Verhalten: {DUPLICATE_POLICY}",
        "",
        "ZUSAMMENFASSUNG",
        "-" * 60,
        f"  Zeilen in Ziel (empty):     {sum(len(r) for r in idx_t.values())}",
        f"  Zeilen in Quelle (Freeze):  {sum(len(r) for r in idx_s.values())}",
        f"  Zeilen uebernommen:         {len(assignments)}",
        f"  Zellen geschrieben:         {cells_written}",
        f"  Zeilen wegen Duplikat uebersprungen: {skipped_dups}",
        f"  Formeln im Ziel durch Werte ersetzt: {formulas_overwritten}",
        "",
    ]
    log += block("NICHT GEMATCHT - nur in 'empty' (Ziel) vorhanden", unmatched_t, "empty")
    log += block("NICHT GEMATCHT - nur in 'Trafo Freeze-Liste' (Quelle) vorhanden", unmatched_s, "Freeze")
    log += block("DOPPELTE KEYS in 'empty' (Ziel)", dup_t, "empty")
    log += block("DOPPELTE KEYS in 'Trafo Freeze-Liste' (Quelle)", dup_s, "Freeze")
    if header_diffs:
        log += [f"UNTERSCHIEDLICHE KOPFZEILEN in kopierten Spalten ({len(header_diffs)})", "-" * 60]
        log += [f"  {d}" for d in header_diffs] + [""]
    if shared_masters:
        log += ["WARNUNG: Ueberschriebene Zellen waren Master einer 'Shared Formula' - abhaengige Formelzellen",
                "ausserhalb des kopierten Bereichs koennen dadurch kaputt sein: " + ", ".join(shared_masters), ""]

    log_path.write_text("\n".join(log), encoding="utf-8")

    print(f"Fertig: {len(assignments)} Zeilen uebernommen, "
          f"{len(unmatched_t)} nur in Ziel, {len(unmatched_s)} nur in Quelle, "
          f"{len(dup_t)} / {len(dup_s)} doppelte Keys (Ziel / Quelle).")
    if header_diffs:
        print(f"Warnung: {len(header_diffs)} abweichende Kopfzeilen-Zellen (siehe Protokoll).")
    if shared_masters:
        print("Warnung: Shared-Formula-Master ueberschrieben (siehe Protokoll).")
    print(f"Ausgabe:   {filled_path}")
    print(f"Protokoll: {log_path}")


if __name__ == "__main__":
    main()
