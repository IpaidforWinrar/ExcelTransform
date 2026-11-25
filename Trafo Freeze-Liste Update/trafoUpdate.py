"""
trafoUpdate.py

Uebernimmt die Werte aus "Trafo Freeze-Liste.xlsx" in eine Kopie der
"empty Trafo Freeze-Liste.xlsx". Zugeordnet wird ueber den zusammengesetzten
Primary Key (Stellennummer + Personalnummer), getrennt fuer jeden Reiter, der
unten in SHEETS eingetragen ist.

Die Zielliste wird NICHT mit openpyxl neu gespeichert: Die Werte werden direkt in
die Sheet-XML der kopierten Datei geschrieben. Alle anderen Teile der Datei
(Datenverbindungen, Power Query, Slicer, Formatierung, Blattschutz, ...) bleiben
unveraendert, ebenso alle Zeilen, in die nichts geschrieben wird. openpyxl wird
nur zum Lesen der Keys und der Quellwerte benutzt.

Ordner (relativ zu diesem Skript):
    update Eingabe/   empty Trafo Freeze-Liste.xlsx
                      Trafo Freeze-Liste.xlsx
    update Ausgabe/   filled Trafo Freeze-Liste.xlsx   (wird erzeugt)
                      trafoUpdate_Protokoll.txt         (wird erzeugt, alle Reiter)

Benoetigt nur openpyxl (sonst Python-Standardbibliothek).
"""

import datetime
import os
import re
import shutil
import sys
import xml.etree.ElementTree as ET
import zipfile
from collections import defaultdict
from decimal import Decimal
from pathlib import Path, PurePosixPath
from xml.sax.saxutils import escape

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

# Reiter, die bearbeitet werden (in dieser Reihenfolge). Pro Reiter:
#   name            Reitername (exakt, auch Sonderzeichen wie & )
#   copy_columns    zu uebernehmende Spalten. Einzelne Spalten und/oder Bereiche,
#                   z.B. ["I:BJ"]  oder  ["I:M", "P", "R:BJ"]
# Optional pro Reiter (sonst gelten die Standardwerte unten):
#   key_columns     Primary-Key-Spalten (zusammengesetzt)
#   exclude_columns Spalten, die trotzdem NICHT uebernommen werden, z.B. ["K", "AA:AC"]
#   first_data_row  erste Datenzeile (darueber sind Kopfzeilen)
SHEETS = [
    {"name": "SOLL Wertstromzuordnung", "copy_columns": ["I:BJ"]},
    {"name": "IST ProdukteServices&Rolle", "copy_columns": ["I:BB"]},
]

# Standardwerte fuer alle Reiter
KEY_COLUMNS = ["D", "F"]   # Stellennummer, Personalnummer
FIRST_DATA_ROW = 3

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


def fmt_columns(cols):
    """[9, 10, 11, 16] -> 'I:K, P'"""
    parts, start = [], None
    for i, c in enumerate(cols):
        if start is None:
            start = c
        if i + 1 == len(cols) or cols[i + 1] != c + 1:
            a, b = get_column_letter(start), get_column_letter(c)
            parts.append(a if start == c else f"{a}:{b}")
            start = None
    return ", ".join(parts)


def norm(value):
    """Key-Wert vereinheitlichen: None -> '', 123.0 -> '123', Leerzeichen entfernen."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = str(value).strip()
    return text.lower() if KEY_IGNORE_CASE else text


def read_keys(ws, key_cols, first_row):
    """Liefert {key_tuple: [zeilennummern]} ab first_row. Komplett leere Keys werden uebersprungen."""
    index = defaultdict(list)
    for row in range(first_row, ws.max_row + 1):
        key = tuple(norm(ws.cell(row, c).value) for c in key_cols)
        if all(part == "" for part in key):
            continue
        index[key].append(row)
    return index


def fmt_key(key):
    return " | ".join(key) if len(key) > 1 else key[0]


# ----------------------------------------------------------------------------
# XML-Ebene (reine Textbearbeitung: alles, was nicht angefasst wird, bleibt byte-genau)
# ----------------------------------------------------------------------------

SHEETDATA_RE = re.compile(r"<sheetData\s*/>|<sheetData(?=[\s>])[^>]*>(.*?)</sheetData>", re.S)
ROW_RE = re.compile(r"<row(?=[\s/>])[^>]*?/>|<row(?=[\s/>])[^>]*>.*?</row>", re.S)
CELL_RE = re.compile(r"<c(?=[\s/>])[^>]*?/>|<c(?=[\s/>])[^>]*>.*?</c>", re.S)
ROW_OPEN_RE = re.compile(r"<row((?:\s[^>]*?)?)(/?)>")
CELL_OPEN_RE = re.compile(r"<c((?:\s[^>]*?)?)(/?)>")
ATTR_RE = re.compile(r"""([\w:.\-]+)\s*=\s*("[^"]*"|'[^']*')""")
ROWNUM_RE = re.compile(r"""<row[^>]*?\sr\s*=\s*["'](\d+)["']""")
CELLREF_RE = re.compile(r"""\sr\s*=\s*["']([A-Z]+)(\d+)["']""")
SPANS_RE = re.compile(r"""\sspans\s*=\s*("[^"]*"|'[^']*')""")
FORMULA_TAG_RE = re.compile(r"<f(?=[\s/>])[^>]*>")
_ILLEGAL_XML = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f]")


def find_sheet_part(zin, sheet_name):
    """Pfad der Sheet-XML im ZIP fuer den Reiternamen (ueber workbook.xml + rels)."""
    wb = ET.fromstring(zin.read("xl/workbook.xml"))
    rid = None
    for sheet in wb.iter(f"{{{NS}}}sheet"):
        if sheet.get("name") == sheet_name:
            rid = sheet.get(f"{{{NS_R}}}id")
    if rid is None:
        raise KeyError(f"Reiter '{sheet_name}' nicht in workbook.xml gefunden")
    rels = ET.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
    for rel in rels:
        if rel.get("Id") == rid:
            target = rel.get("Target")
            if target.startswith("/"):
                return target.lstrip("/")
            return str(PurePosixPath("xl") / target)
    raise KeyError(f"Relationship {rid} nicht gefunden")


def cell_xml(ref, attrs, src_cell, epoch):
    """Neues <c>-Element. Alle Attribute des alten Elements (Stil s=...) bleiben, nur t wird neu gesetzt."""
    kept = " ".join(f"{n}={v}" for n, v in attrs if n not in ("r", "t"))
    prefix = f'<c r="{ref}"' + (f" {kept}" if kept else "")

    v = src_cell.value
    if v is None:
        return prefix + "/>"
    if src_cell.data_type == "e":
        return f'{prefix} t="e"><v>{escape(str(v))}</v></c>'
    if isinstance(v, bool):
        return f'{prefix} t="b"><v>{1 if v else 0}</v></c>'
    if isinstance(v, (int, Decimal)):
        return f"{prefix}><v>{v}</v></c>"
    if isinstance(v, float):
        return f"{prefix}><v>{v!r}</v></c>"
    if isinstance(v, (datetime.datetime, datetime.date, datetime.time, datetime.timedelta)):
        return f"{prefix}><v>{to_excel(v, epoch=epoch)!r}</v></c>"
    text = escape(_ILLEGAL_XML.sub("", str(v)))
    return f'{prefix} t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'


def patch_row(row_text, r, cols_src, epoch, stats):
    """Zeile r: Zellen aus cols_src ({spalte: quellzelle}) schreiben, Rest der Zeile bleibt unveraendert."""
    m = ROW_OPEN_RE.match(row_text)
    open_attrs = SPANS_RE.sub("", m.group(1))  # spans ist nur ein Optimierungshinweis
    inner = "" if m.group(2) else row_text[m.end():-len("</row>")]

    cell_ms = list(CELL_RE.finditer(inner))
    by_col = {column_index_from_string(CELLREF_RE.search(c.group()).group(1)): c.group() for c in cell_ms}
    head = inner[:cell_ms[0].start()] if cell_ms else inner
    tail = inner[cell_ms[-1].end():] if cell_ms else ""

    for c, src_cell in cols_src.items():
        ref = f"{get_column_letter(c)}{r}"
        old = by_col.get(c)
        attrs = []
        if old is not None:
            om = CELL_OPEN_RE.match(old)
            attrs = ATTR_RE.findall(om.group(1))
            ftag = FORMULA_TAG_RE.search(old[om.end():])
            if ftag:
                stats["formulas"] += 1
                if re.search(r"""\bt\s*=\s*["']shared["']""", ftag.group()) and re.search(r"\bref\s*=", ftag.group()):
                    stats["shared_masters"].append(ref)
        by_col[c] = cell_xml(ref, attrs, src_cell, epoch)
        stats["cells"] += 1

    body = "".join(by_col[c] for c in sorted(by_col))
    return f"<row{open_attrs}>{head if not cell_ms else ''}{body}{tail}</row>"


def patch_sheet_xml(xml, plan, epoch):
    """plan: {zeilennummer: {spalte: quellzelle}}. Gibt (neues_xml, stats) zurueck."""
    stats = {"cells": 0, "formulas": 0, "shared_masters": []}
    sd = SHEETDATA_RE.search(xml)
    if sd is None:
        raise ValueError("<sheetData> nicht gefunden (ungewoehnliches Sheet-XML-Format)")
    body = sd.group(1) or ""
    rows = {}
    for m in ROW_RE.finditer(body):
        rows[int(ROWNUM_RE.search(m.group()).group(1))] = m.group()

    for r, cols_src in plan.items():
        rows[r] = patch_row(rows.get(r, f'<row r="{r}"/>'), r, cols_src, epoch, stats)

    new_body = "".join(rows[r] for r in sorted(rows))
    return xml[:sd.start()] + f"<sheetData>{new_body}</sheetData>" + xml[sd.end():], stats


def set_full_calc(wb_text):
    """calcPr fullCalcOnLoad="1" setzen. None, wenn workbook.xml ungewoehnlich aufgebaut ist."""
    m = re.search(r"<calcPr(?=[\s/>])[^>]*?/?>", wb_text)
    if m:
        tag = m.group()
        if "fullCalcOnLoad" in tag:
            new = re.sub(r"""fullCalcOnLoad\s*=\s*("[^"]*"|'[^']*')""", 'fullCalcOnLoad="1"', tag)
        else:
            new = re.sub(r"\s*(/?)>$", r' fullCalcOnLoad="1"\1>', tag)
        return wb_text[:m.start()] + new + wb_text[m.end():]
    successors = re.search(
        r"<(oleSize|customWorkbookViews|pivotCaches|smartTagPr|smartTagTypes|webPublishing|"
        r"fileRecoveryPr|webPublishObjects|extLst)(?=[\s/>])", wb_text)
    pos = successors.start() if successors else wb_text.rfind("</workbook>")
    if pos < 0:
        return None
    return wb_text[:pos] + '<calcPr fullCalcOnLoad="1"/>' + wb_text[pos:]


def drop_calc_chain(zin):
    """calcChain.xml aus Content-Types und Beziehungen austragen (Excel baut sie neu auf)."""
    ct = zin.read("[Content_Types].xml").decode("utf-8")
    ct = re.sub(r"""<Override\b[^>]*PartName\s*=\s*["']/xl/calcChain\.xml["'][^>]*/>""", "", ct)
    rels = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rels = re.sub(r"""<Relationship\b[^>]*/calcChain["'][^>]*/>""", "", rels)
    return {"[Content_Types].xml": ct.encode("utf-8"), "xl/_rels/workbook.xml.rels": rels.encode("utf-8")}


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
# Matching pro Reiter
# ----------------------------------------------------------------------------

def match_sheet(cfg, wb_target, wb_source):
    """Keys vergleichen und festlegen, welche Zielzeile welche Quellzeile bekommt (noch nichts schreiben)."""
    name = cfg["name"]
    first_row = cfg.get("first_data_row", FIRST_DATA_ROW)
    key_cols = parse_columns(cfg.get("key_columns", KEY_COLUMNS))
    excluded = set(parse_columns(cfg.get("exclude_columns", [])))
    copy_cols = [c for c in parse_columns(cfg["copy_columns"]) if c not in excluded]
    if not copy_cols:
        sys.exit(f"Reiter '{name}': keine Spalten zum Kopieren konfiguriert (copy_columns).")
    ws_t, ws_s = wb_target[name], wb_source[name]

    header_diffs = []
    for row in range(1, first_row):
        for c in copy_cols:
            a, b = norm(ws_t.cell(row, c).value), norm(ws_s.cell(row, c).value)
            if a != b:
                header_diffs.append(f"{get_column_letter(c)}{row}: Ziel='{a}' / Quelle='{b}'")

    idx_t = read_keys(ws_t, key_cols, first_row)
    idx_s = read_keys(ws_s, key_cols, first_row)

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

    plan = {
        tgt_row: {c: ws_s.cell(src_row, c) for c in copy_cols
                  if not isinstance(ws_t.cell(tgt_row, c), MergedCell)}
        for tgt_row, src_row in assignments.items()
    }
    return {
        "name": name, "key_cols": key_cols, "copy_cols": copy_cols,
        "idx_t": idx_t, "idx_s": idx_s,
        "dup_t": {k: r for k, r in idx_t.items() if len(r) > 1},
        "dup_s": {k: r for k, r in idx_s.items() if len(r) > 1},
        "unmatched_t": {k: r for k, r in idx_t.items() if k not in idx_s},
        "unmatched_s": {k: r for k, r in idx_s.items() if k not in idx_t},
        "assignments": assignments, "skipped_dups": skipped_dups,
        "header_diffs": header_diffs, "plan": plan, "stats": None,
    }


def sheet_log(res):
    """Protokollzeilen fuer einen Reiter."""
    def block(title, mapping, label):
        lines = [f"{title} ({len(mapping)})", "-" * 60]
        if not mapping:
            lines.append("  (keine)")
        for key, rows_ in sorted(mapping.items()):
            lines.append(f"  {fmt_key(key)}   [{label} Zeile {', '.join(map(str, rows_))}]")
        lines.append("")
        return lines

    stats = res["stats"]
    title = f"REITER: {res['name']}"
    log = [
        "=" * 70, title, "=" * 70,
        f"Primary Key (Spalten): {' + '.join(get_column_letter(c) for c in res['key_cols'])}"
        "  - Format unten: Stellennummer | Personalnummer",
        f"Kopierte Spalten: {fmt_columns(res['copy_cols'])}",
        "",
        "ZUSAMMENFASSUNG",
        "-" * 60,
        f"  Zeilen in Ziel (empty):     {sum(len(r) for r in res['idx_t'].values())}",
        f"  Zeilen in Quelle (Freeze):  {sum(len(r) for r in res['idx_s'].values())}",
        f"  Zeilen uebernommen:         {len(res['assignments'])}",
        f"  Zellen geschrieben:         {stats['cells']}",
        f"  Zeilen wegen Duplikat uebersprungen: {res['skipped_dups']}",
        f"  Formeln im Ziel durch Werte ersetzt: {stats['formulas']}",
        "",
    ]
    log += block("NICHT GEMATCHT - nur in 'empty' (Ziel) vorhanden", res["unmatched_t"], "empty")
    log += block("NICHT GEMATCHT - nur in 'Trafo Freeze-Liste' (Quelle) vorhanden", res["unmatched_s"], "Freeze")
    log += block("DOPPELTE KEYS in 'empty' (Ziel)", res["dup_t"], "empty")
    log += block("DOPPELTE KEYS in 'Trafo Freeze-Liste' (Quelle)", res["dup_s"], "Freeze")
    if res["header_diffs"]:
        log += [f"UNTERSCHIEDLICHE KOPFZEILEN in kopierten Spalten ({len(res['header_diffs'])})", "-" * 60]
        log += [f"  {d}" for d in res["header_diffs"]] + [""]
    if stats["shared_masters"]:
        log += ["WARNUNG: Ueberschriebene Zellen waren Master einer 'Shared Formula' - abhaengige Formelzellen",
                "ausserhalb des kopierten Bereichs koennen dadurch kaputt sein: " + ", ".join(stats["shared_masters"]), ""]
    return log


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

    # 1) empty -> Ausgabe kopieren und umbenennen
    out_dir.mkdir(exist_ok=True)
    shutil.copyfile(empty_path, filled_path)
    print(f"Kopiert: {empty_path.name} -> {filled_path}")

    # 2) Nur LESEN mit openpyxl (wird nie gespeichert): Keys und Quellwerte
    wb_target = load_workbook(filled_path, data_only=True)
    wb_source = load_workbook(source_path, data_only=True)
    for cfg in SHEETS:
        for label, wb in (("Ziel", wb_target), ("Quelle", wb_source)):
            if cfg["name"] not in wb.sheetnames:
                sys.exit(f"Reiter '{cfg['name']}' fehlt in {label}-Datei. Vorhanden: {wb.sheetnames}")

    # 3) Pro Reiter: Keys vergleichen
    results = [match_sheet(cfg, wb_target, wb_source) for cfg in SHEETS]

    # 4) Werte direkt in die Sheet-XML der kopierten Datei schreiben (alle Reiter in einem Durchgang)
    recalc_failed = False
    with zipfile.ZipFile(filled_path) as zin:
        replacements, removed = {}, set()
        for res in results:
            part = find_sheet_part(zin, res["name"])
            new_xml, res["stats"] = patch_sheet_xml(zin.read(part).decode("utf-8"), res["plan"], wb_source.epoch)
            replacements[part] = new_xml.encode("utf-8")
        if RECALC_ON_OPEN and any(res["assignments"] for res in results):
            wb_xml = set_full_calc(zin.read("xl/workbook.xml").decode("utf-8"))
            if wb_xml is None:
                recalc_failed = True
            else:
                replacements["xl/workbook.xml"] = wb_xml.encode("utf-8")
        if any(res["stats"]["formulas"] for res in results) and "xl/calcChain.xml" in zin.namelist():
            replacements.update(drop_calc_chain(zin))
            removed.add("xl/calcChain.xml")

    rewrite_zip(filled_path, replacements, removed)

    # 5) Protokoll (eine Datei, ein Abschnitt pro Reiter)
    log = [
        f"trafoUpdate Protokoll - {datetime.datetime.now():%d.%m.%Y %H:%M:%S}",
        f"Bearbeitete Reiter: {', '.join(r['name'] for r in results)}",
        f"Duplikat-Verhalten: {DUPLICATE_POLICY}",
        "",
    ]
    for res in results:
        log += sheet_log(res)
    if recalc_failed:
        log += ["HINWEIS: 'Neuberechnen beim Oeffnen' konnte nicht gesetzt werden (workbook.xml ungewoehnlich).",
                "In Excel einmal Strg+Alt+Umschalt+F9 druecken.", ""]
    log_path.write_text("\n".join(log), encoding="utf-8")

    for res in results:
        print(f"\n[{res['name']}] {len(res['assignments'])} Zeilen uebernommen, "
              f"{len(res['unmatched_t'])} nur in Ziel, {len(res['unmatched_s'])} nur in Quelle, "
              f"{len(res['dup_t'])} / {len(res['dup_s'])} doppelte Keys (Ziel / Quelle).")
        if res["header_diffs"]:
            print(f"  Warnung: {len(res['header_diffs'])} abweichende Kopfzeilen-Zellen (siehe Protokoll).")
        if res["stats"]["shared_masters"]:
            print("  Warnung: Shared-Formula-Master ueberschrieben (siehe Protokoll).")
    print(f"\nAusgabe:   {filled_path}")
    print(f"Protokoll: {log_path}")


if __name__ == "__main__":
    main()
