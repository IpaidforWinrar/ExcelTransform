"""
Erstellt "Aktenplan bearbeitet.xlsx" aus dem neuen Aktenplan und vergleicht ihn mit dem alten.

Ablauf:
1. "AktenplanNeu.xlsx" (Spalte A bis Q, Werte ab Zeile 3) wird in "empty Aktenplan.xlsx" übertragen,
   ebenfalls ab Zeile 3, aber zwei Spalten nach rechts: Spalte A -> C, B -> D, ... Q -> S.
   Zeile 1 der Vorlage bekommt die Überschrift
   "Berechtigung im Aktenplan ZB IT - Stand <Monat> <Jahr>".
2. Neu-Markierung: In beiden Tabellen bilden Spalte D und E einen eindeutigen Schlüssel.
   Zeilen aus "Aktenplan bearbeitet.xlsx", deren Schlüssel in "AktenplanAlt.xlsx" fehlt, bekommen
   in Spalte A das Wort "neu" in roter Schrift (ab Zeile 3, Zeile 2 ist die Überschrift).
3. Bereich eintragen: Jeder Wert aus Spalte C wird in Spalte A von "Aktenplan Zuordnung.xlsx"
   gesucht. Bei einem Treffer kommt der Wert aus Spalte B der Zuordnung in Spalte B der Zeile.
   Zuvor wird der Inhalt von Spalte B gelesen: Ist die Zelle leer oder gleich, ist alles gut.
   Steht dort ein anderer Wert, wird er überschrieben und in der Logdatei festgehalten.

Eingabe-Dateien (AktenplanNeu.xlsx, AktenplanAlt.xlsx, empty Aktenplan.xlsx, Aktenplan Zuordnung.xlsx):
    FAVIS DB Aktenplan/Eingabe Aktenplaene
    Dieser Ordner ist schreibgeschützt: Das Skript liest dort nur. "empty Aktenplan.xlsx" wird in den
    Ausgabe-Ordner kopiert und als "Aktenplan bearbeitet.xlsx" bearbeitet.
Ausgabe ("Aktenplan bearbeitet.xlsx" und "Aktenplan Auswertung.log" mit allem, was im Terminal steht):
    FAVIS DB Aktenplan/Ausgabe Aktenplan

Aufruf:
    python aktenplanAuswertung.py
"""

import os
import re
import shutil
import stat
import sys
import traceback
from copy import copy
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font

# ORDNERLOGIK
ORDNER = Path(__file__).resolve().parent
EINGABE_ORDNER = ORDNER.parent / "FAVIS DB Aktenplan" / "Eingabe Aktenplaene"
AUSGABE_ORDNER = ORDNER.parent / "FAVIS DB Aktenplan" / "Ausgabe Aktenplan"

NEU_DATEI = EINGABE_ORDNER / "AktenplanNeu.xlsx"
ALT_DATEI = EINGABE_ORDNER / "AktenplanAlt.xlsx"
VORLAGE_DATEI = EINGABE_ORDNER / "empty Aktenplan.xlsx"
ZUORDNUNG_DATEI = EINGABE_ORDNER / "Aktenplan Zuordnung.xlsx"
AUSGABE_DATEI = AUSGABE_ORDNER / "Aktenplan bearbeitet.xlsx"
LOG_DATEI = AUSGABE_ORDNER / "Aktenplan Auswertung.log"

START_ROW = 3  # Zeile 1 = Info, Zeile 2 = Spaltenüberschriften
NEU_SPALTEN = 17  # Spalte A bis Q der neuen Tabelle
SPALTEN_VERSATZ = 2  # A -> C

COL_MARKIERUNG = 1  # Spalte A
COL_BEREICH = 2  # Spalte B
COL_EINTRAG = 3  # Spalte C
COL_KEY1 = 4  # Spalte D
COL_KEY2 = 5  # Spalte E

MONATE = [
    "Januar", "Februar", "März", "April", "Mai", "Juni",
    "Juli", "August", "September", "Oktober", "November", "Dezember",
]


def clean(value) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def schluessel(zeile) -> tuple:
    return (clean(zeile[COL_KEY1 - 1]).casefold(), clean(zeile[COL_KEY2 - 1]).casefold())


def pruefe_dateien(*pfade: Path) -> None:
    fehlend = [p for p in pfade if not p.exists()]
    if fehlend:
        sys.exit("Datei nicht gefunden: " + ", ".join(str(p) for p in fehlend))


def lese_alt_schluessel(path: Path) -> set:
    ws = load_workbook(path, read_only=True, data_only=True).active
    keys = set()
    for zeile in ws.iter_rows(min_row=START_ROW, max_col=COL_KEY2, values_only=True):
        key = schluessel(zeile)
        if key != ("", ""):
            keys.add(key)
    return keys


def lese_zuordnung(path: Path):
    """Liefert {eintrag: wert} aus Spalte A und B von "Aktenplan Zuordnung.xlsx" (ab Zeile 2)."""
    ws = load_workbook(path, read_only=True, data_only=True).active
    zeilen = list(ws.iter_rows(values_only=True))
    if not zeilen:
        sys.exit(f"{path.name} ist leer.")

    kopf = [clean(c).casefold() for c in zeilen[0]]
    wert_col = 1  # Spalte B
    status_col = kopf.index("status") if "status" in kopf else None
    if len(kopf) > 1 and kopf[1] == "status" and "bereiche" in kopf:
        # Ausgabe von aktenplan.py: Spalte B enthält nur unique/mehrfach, der Bereich steht in "Bereiche"
        wert_col = kopf.index("bereiche")
        print(f"Hinweis: Spalte B von {path.name} ist 'Status', der Bereich wird aus der Spalte "
              f"'Bereiche' ({chr(ord('A') + wert_col)}) gelesen.")

    zuordnung = {}
    mehrfach = set()
    for zeile in zeilen[1:]:
        eintrag = clean(zeile[0] if zeile else None)
        wert = clean(zeile[wert_col] if len(zeile) > wert_col else None)
        wert = re.sub(r"\s*\(\d+x\)", "", wert)  # "IT 2 (2x)" -> "IT 2"
        if not eintrag or not wert:
            continue
        key = eintrag.casefold()
        zuordnung.setdefault(key, wert)
        if status_col is not None and clean(zeile[status_col]).casefold() == "mehrfach":
            mehrfach.add(key)
    return zuordnung, mehrfach


def main() -> None:
    pruefe_dateien(NEU_DATEI, ALT_DATEI, VORLAGE_DATEI, ZUORDNUNG_DATEI)

    # Schritt 1: neuen Aktenplan in die Vorlage übertragen
    neu = load_workbook(NEU_DATEI, data_only=True).active
    # Der Eingabe-Ordner ist schreibgeschützt: Die Vorlage wird als "Aktenplan bearbeitet.xlsx"
    # in den Ausgabe-Ordner kopiert, alle weiteren Änderungen passieren nur an dieser Kopie.
    AUSGABE_ORDNER.mkdir(parents=True, exist_ok=True)
    if AUSGABE_DATEI.exists():
        os.chmod(AUSGABE_DATEI, stat.S_IWRITE)  # Ergebnis eines früheren Laufs darf überschrieben werden
    shutil.copyfile(VORLAGE_DATEI, AUSGABE_DATEI)  # ohne Dateirechte, die Kopie ist beschreibbar
    os.chmod(AUSGABE_DATEI, stat.S_IWRITE)
    wb = load_workbook(AUSGABE_DATEI)
    ws = wb.active

    heute = datetime.now()
    ws.cell(row=1, column=1, value=f"Berechtigung im Aktenplan ZB IT - Stand {MONATE[heute.month - 1]} {heute.year}")

    letzte_zeile = START_ROW - 1
    for row_idx, zeile in enumerate(
        neu.iter_rows(min_row=START_ROW, max_col=NEU_SPALTEN), start=START_ROW
    ):
        if all(clean(c.value) == "" for c in zeile):
            continue
        letzte_zeile = row_idx
        for col_idx, quelle in enumerate(zeile, start=1):
            # Nur Werte übertragen: Formate, Schriften und leere Zellen der Vorlage bleiben unverändert
            if quelle.value is not None:
                ws.cell(row=row_idx, column=col_idx + SPALTEN_VERSATZ).value = quelle.value
    print(f"{letzte_zeile - START_ROW + 1} Zeilen aus {NEU_DATEI.name} übertragen (bis Zeile {letzte_zeile}).")

    # Schritt 2: Zeilen markieren, die es im alten Aktenplan nicht gab
    alt_keys = lese_alt_schluessel(ALT_DATEI)
    anzahl_neu = 0
    for row_idx in range(START_ROW, letzte_zeile + 1):
        zeile = [ws.cell(row=row_idx, column=c).value for c in range(1, COL_KEY2 + 1)]
        key = schluessel(zeile)
        if key == ("", "") or key in alt_keys:
            continue
        zelle = ws.cell(row=row_idx, column=COL_MARKIERUNG, value="neu")
        schrift = copy(zelle.font)
        schrift.color = "FFFF0000"
        zelle.font = schrift
        anzahl_neu += 1
    print(f"Neu gegenüber {ALT_DATEI.name}: {anzahl_neu}")

    # Schritt 3: Bereich aus der Zuordnung in Spalte B eintragen
    zuordnung, mehrfach = lese_zuordnung(ZUORDNUNG_DATEI)
    eingetragen = 0
    ohne_treffer = 0
    abweichungen = []
    mehrfach_zeilen = []
    for row_idx in range(START_ROW, letzte_zeile + 1):
        eintrag = clean(ws.cell(row=row_idx, column=COL_EINTRAG).value)
        if not eintrag:
            continue
        wert = zuordnung.get(eintrag.casefold())
        if wert is None:
            ohne_treffer += 1
            continue

        zelle = ws.cell(row=row_idx, column=COL_BEREICH)
        vorher = clean(zelle.value)
        if vorher and vorher != wert:
            abweichungen.append((row_idx, eintrag, vorher, wert))
        zelle.value = wert
        eingetragen += 1
        if eintrag.casefold() in mehrfach:
            mehrfach_zeilen.append((row_idx, eintrag, wert))

    print(f"Bereich eingetragen: {eingetragen}, ohne Treffer in {ZUORDNUNG_DATEI.name}: {ohne_treffer}")
    if mehrfach_zeilen:
        print(f"Hinweis: {len(mehrfach_zeilen)} Zeilen mit mehrfach zugeordnetem Eintrag (mehrere Bereiche):")
        for row_idx, eintrag, wert in mehrfach_zeilen:
            print(f"  Zeile {row_idx}: {eintrag} -> {wert}")
    if abweichungen:
        print(f"Abweichung: {len(abweichungen)} Zellen in Spalte B hatten einen anderen Wert und wurden überschrieben:")
        for row_idx, eintrag, vorher, wert in abweichungen:
            print(f"  Zeile {row_idx}: {eintrag}: vorher '{vorher}', jetzt '{wert}'")
    else:
        print("Keine Abweichungen in Spalte B.")

    wb.save(AUSGABE_DATEI)
    print(f"Fertig: {AUSGABE_DATEI}")
    print(f"Log: {LOG_DATEI}")


class Tee:
    """Schreibt alles gleichzeitig ins Terminal und in die Logdatei."""

    def __init__(self, *streams):
        self.streams = streams

    def write(self, text):
        for stream in self.streams:
            stream.write(text)

    def flush(self):
        for stream in self.streams:
            stream.flush()


def main_mit_log() -> None:
    AUSGABE_ORDNER.mkdir(parents=True, exist_ok=True)
    with open(LOG_DATEI, "w", encoding="utf-8") as log:
        stdout, stderr = sys.stdout, sys.stderr
        sys.stdout, sys.stderr = Tee(stdout, log), Tee(stderr, log)
        try:
            print(f"Lauf vom {datetime.now():%d.%m.%Y %H:%M:%S}")
            main()
        except SystemExit as exc:
            if isinstance(exc.code, str):
                print(exc.code, file=sys.stderr)
                raise SystemExit(1)
            raise
        except Exception:
            traceback.print_exc()
            raise SystemExit(1)
        finally:
            sys.stdout, sys.stderr = stdout, stderr


if __name__ == "__main__":
    main_mit_log()
