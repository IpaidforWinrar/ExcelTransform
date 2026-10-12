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

## Zuordnungspaare (zuordnung.py)

Vergleicht Spalte A (IT 1 bis IT 6, "IT-Leitung + Stab", "übergreifend") mit Spalte B ab Zeile 3
und gibt eine Exceltabelle mit einer Spalte pro Bereich aus. Die Extraspalte "doppelt" listet
Begriffe, die mehrfach vorkommen, mit ihren Bereichen, z. B. `Gruppe X → IT 1, IT 3`.

```bash
python zuordnung.py
```

Ohne Argument wird `test.xlsx` aus dem Ordner des Skripts gelesen und das Ergebnis als
`test_Zuordnung.xlsx` gespeichert. Eine andere Datei kann als Argument angegeben werden.
