"""
Bearbeitet die Exceltabelle "Auswertung Benutzergruppen.xlsx":

1. Löscht die ersten 3 Zeilen.
2. Merged Spalte B in Spalte C: Ist die Zelle in C leer, wird der Wert aus B übernommen,
   sonst bleibt C unverändert.
3. Löscht Spalte B, dadurch wird die bisherige Spalte C zur neuen Spalte B.

Aufruf:
    python transform.py [Eingabedatei] [Ausgabedatei]

Ohne Argumente wird "Auswertung Benutzergruppen.xlsx" im aktuellen Ordner verwendet
und das Ergebnis als "Auswertung Benutzergruppen_bearbeitet.xlsx" gespeichert.
"""

import sys
from pathlib import Path

from openpyxl import load_workbook

# ORDNERLOGIK
# relative Ordnerstruktur ziehen & anwenden
ORDNER = Path(__file__).resolve().parent
DEFAULT_INPUT = ORDNER.parent / "FAVIS DB Benutzergruppen" / "Eingabe Benutzergruppen" / "Auswertung Benutzergruppen.xlsx"
DEFAULT_OUTPUT = ORDNER.parent / "FAVIS DB Benutzergruppen" / "Ausgabe Benutzergruppen"


COL_B = 2
COL_C = 3


def is_empty(value) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def transform(input_path: Path, output_path: Path) -> None:
    wb = load_workbook(input_path)
    ws = wb.active

    # Schritt 1: die ersten 3 Zeilen löschen
    ws.delete_rows(1, 3)

    # Schritt 2: Spalte B in Spalte C mergen
    for row in range(1, ws.max_row + 1):
        cell_b = ws.cell(row=row, column=COL_B)
        cell_c = ws.cell(row=row, column=COL_C)
        if is_empty(cell_c.value):
            cell_c.value = cell_b.value

    # Schritt 3: Spalte B löschen -> Spalte C wird zu Spalte B
    ws.delete_cols(COL_B)

    wb.save(output_path)


def main() -> None:
    input_path = (
        Path(sys.argv[1])
        if len(sys.argv) > 1
        else DEFAULT_INPUT
    )

    if len(sys.argv) > 2:
        output_path = Path(sys.argv[2])
    else:
        output_path = (
                DEFAULT_OUTPUT
                / f"{input_path.stem}_bearbeitet{input_path.suffix}"
        )

    if not input_path.exists():
        sys.exit(f"Datei nicht gefunden: {input_path}")

    output_path.parent.mkdir(parents=True, exist_ok=True)

    transform(input_path, output_path)
    print(f"Fertig: {output_path}")


if __name__ == "__main__":
    main()
