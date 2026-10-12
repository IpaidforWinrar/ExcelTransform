# ExcelTransform

Bearbeitet die Exceltabelle "Auswertung Benutzergruppen.xlsx" mit openpyxl:

1. Löscht die ersten 3 Zeilen.
2. Merged Spalte B in Spalte C: Ist die Zelle in C leer, wird der Wert aus B übernommen.
3. Löscht Spalte B, dadurch wird Spalte C zur neuen Spalte B.

## Installation

```bash
pip install -r requirements.txt
```

## Aufruf

```bash
python transform.py "Auswertung Benutzergruppen.xlsx"
```

Das Ergebnis wird als `Auswertung Benutzergruppen_bearbeitet.xlsx` gespeichert. Optional kann als zweites Argument ein eigener Name für die Ausgabedatei angegeben werden.
