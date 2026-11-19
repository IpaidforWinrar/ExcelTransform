"""
trafoUpdate.py

Uebernimmt die Werte aus "Trafo Freeze-Liste.xlsx" in eine Kopie der
"empty Trafo Freeze-Liste.xlsx". Zugeordnet wird ueber den zusammengesetzten
Primary Key (Stellennummer + Personalnummer) im Reiter "SOLL Wertstromzuordnung".

Ordner (relativ zu diesem Skript):
    update Eingabe/   empty Trafo Freeze-Liste.xlsx
                      Trafo Freeze-Liste.xlsx
    update Ausgabe/   filled Trafo Freeze-Liste.xlsx   (wird erzeugt)
                      trafoUpdate_Protokoll.txt         (wird erzeugt)

Benoetigt: pip install openpyxl
"""

import shutil
import sys
import zipfile
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import column_index_from_string, get_column_letter

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

# Quelle: True = berechnete Werte uebernehmen (bei Formeln das Ergebnis),
#         False = Formeln als Formeltext uebernehmen
SOURCE_VALUES_ONLY = True

# Was tun, wenn ein Key mehrfach vorkommt (in Quelle oder Ziel)?
#   "skip"  = nicht uebernehmen, nur im Protokoll auflisten
#   "first" = jeweils die erste Zeile verwenden
DUPLICATE_POLICY = "skip"

# Gross-/Kleinschreibung beim Vergleichen der Keys ignorieren?
KEY_IGNORE_CASE = True

# =============================================================================


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


def warn_lost_parts(path):
    """openpyxl verwirft Verbindungen / Power Query beim Speichern -> vorher pruefen."""
    warnings = []
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        if "xl/connections.xml" in names:
            warnings.append("xl/connections.xml (Datenverbindungen)")
        if any(n.startswith("customXml/") for n in names):
            warnings.append("customXml/ (z.B. Power Query / DataMashup)")
        if any(n.startswith("xl/slicers") or n.startswith("xl/timelines") for n in names):
            warnings.append("Slicer / Zeitachsen")
    return warnings


def main():
    base = Path(__file__).resolve().parent
    in_dir, out_dir = base / INPUT_DIR, base / OUTPUT_DIR
    empty_path, source_path = in_dir / EMPTY_FILE, in_dir / SOURCE_FILE
    filled_path, log_path = out_dir / FILLED_FILE, out_dir / LOG_FILE

    for p in (empty_path, source_path):
        if not p.exists():
            sys.exit(f"Datei nicht gefunden: {p}")

    key_cols = parse_columns(KEY_COLUMNS)
    copy_cols = [c for c in parse_columns(COPY_COLUMNS) if c not in parse_columns(EXCLUDE_COLUMNS)]
    if not copy_cols:
        sys.exit("Keine Spalten zum Kopieren konfiguriert (COPY_COLUMNS).")

    # 1) empty -> Ausgabe kopieren und umbenennen
    out_dir.mkdir(exist_ok=True)
    shutil.copyfile(empty_path, filled_path)
    print(f"Kopiert: {empty_path.name} -> {filled_path}")

    lost = warn_lost_parts(filled_path)
    if lost:
        print("\nACHTUNG: Die Datei enthaelt Teile, die openpyxl beim Speichern verwirft:")
        for item in lost:
            print(f"  - {item}")
        print("Abfragen/Verbindungen sind in der Ausgabe danach evtl. nicht mehr vorhanden.\n")

    # 2) Dateien laden
    wb_target = load_workbook(filled_path)
    wb_source = load_workbook(source_path, data_only=SOURCE_VALUES_ONLY)
    for name, wb in (("Ziel", wb_target), ("Quelle", wb_source)):
        if SHEET_NAME not in wb.sheetnames:
            sys.exit(f"Reiter '{SHEET_NAME}' fehlt in {name}-Datei. Vorhanden: {wb.sheetnames}")
    ws_t, ws_s = wb_target[SHEET_NAME], wb_source[SHEET_NAME]

    # Kopfzeilen der zu kopierenden Spalten vergleichen (Plausibilitaetscheck)
    header_diffs = []
    for row in range(1, FIRST_DATA_ROW):
        for c in copy_cols:
            a, b = norm(ws_t.cell(row, c).value), norm(ws_s.cell(row, c).value)
            if a != b:
                header_diffs.append(f"{get_column_letter(c)}{row}: Ziel='{a}' / Quelle='{b}'")

    # 3) Keys einlesen
    idx_t = read_keys(ws_t, key_cols)
    idx_s = read_keys(ws_s, key_cols)

    dup_t = {k: r for k, r in idx_t.items() if len(r) > 1}
    dup_s = {k: r for k, r in idx_s.items() if len(r) > 1}
    unmatched_t = {k: r for k, r in idx_t.items() if k not in idx_s}
    unmatched_s = {k: r for k, r in idx_s.items() if k not in idx_t}

    # 4) Werte uebernehmen
    rows_copied = 0
    skipped_dups = 0
    for key, rows_t in idx_t.items():
        if key not in idx_s:
            continue
        rows_s = idx_s[key]
        if DUPLICATE_POLICY == "skip" and (len(rows_t) > 1 or len(rows_s) > 1):
            skipped_dups += len(rows_t)
            continue
        src_row = rows_s[0]
        for tgt_row in rows_t:
            for c in copy_cols:
                cell = ws_t.cell(tgt_row, c)
                if isinstance(cell, MergedCell):
                    continue
                cell.value = ws_s.cell(src_row, c).value
            rows_copied += 1

    wb_target.save(filled_path)

    # 5) Protokoll
    def block(title, mapping, label):
        lines = [f"{title} ({len(mapping)})", "-" * 60]
        if not mapping:
            lines.append("  (keine)")
        for key, rows in sorted(mapping.items()):
            lines.append(f"  {fmt_key(key)}   [{label} Zeile {', '.join(map(str, rows))}]")
        lines.append("")
        return lines

    key_names = " + ".join(get_column_letter(c) for c in key_cols)
    copy_names = ", ".join(get_column_letter(c) for c in copy_cols)
    log = [
        f"trafoUpdate Protokoll - {datetime.now():%d.%m.%Y %H:%M:%S}",
        f"Reiter: {SHEET_NAME}",
        f"Primary Key (Spalten): {key_names}  - Format im Protokoll: Stellennummer | Personalnummer",
        f"Kopierte Spalten: {copy_names}",
        f"Duplikat-Verhalten: {DUPLICATE_POLICY}",
        "",
        "ZUSAMMENFASSUNG",
        "-" * 60,
        f"  Zeilen in Ziel (empty):     {sum(len(r) for r in idx_t.values())}",
        f"  Zeilen in Quelle (Freeze):  {sum(len(r) for r in idx_s.values())}",
        f"  Zeilen uebernommen:         {rows_copied}",
        f"  Zeilen wegen Duplikat uebersprungen: {skipped_dups}",
        "",
    ]
    log += block("NICHT GEMATCHT - nur in 'empty' (Ziel) vorhanden", unmatched_t, "empty")
    log += block("NICHT GEMATCHT - nur in 'Trafo Freeze-Liste' (Quelle) vorhanden", unmatched_s, "Freeze")
    log += block("DOPPELTE KEYS in 'empty' (Ziel)", dup_t, "empty")
    log += block("DOPPELTE KEYS in 'Trafo Freeze-Liste' (Quelle)", dup_s, "Freeze")
    if header_diffs:
        log += [f"UNTERSCHIEDLICHE KOPFZEILEN in kopierten Spalten ({len(header_diffs)})", "-" * 60]
        log += [f"  {d}" for d in header_diffs] + [""]
    if lost:
        log += ["HINWEIS: Beim Speichern evtl. verworfene Teile: " + "; ".join(lost), ""]

    log_path.write_text("\n".join(log), encoding="utf-8")

    print(f"Fertig: {rows_copied} Zeilen uebernommen, "
          f"{len(unmatched_t)} nur in Ziel, {len(unmatched_s)} nur in Quelle, "
          f"{len(dup_t)} / {len(dup_s)} doppelte Keys (Ziel / Quelle).")
    if header_diffs:
        print(f"Warnung: {len(header_diffs)} abweichende Kopfzeilen-Zellen (siehe Protokoll).")
    print(f"Ausgabe:   {filled_path}")
    print(f"Protokoll: {log_path}")


if __name__ == "__main__":
    main()
