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

## Gruppen zuordnen (matcher.py)

Läuft nach `transform.py`. Benötigt im Ordner des Skripts:

- `Auswertung Benutzergruppen_bearbeitet.xlsx` (Ergebnis von `transform.py`)
- `empty Auswertung Benutzergruppen ZB IT.xlsx` (Vorlage: A = betrifft, B = Gruppe, C = Name, ab Zeile 3)
- `Benutzergruppen Zuordnung.xlsx` (Zeile 1 = Bereiche, darunter die Gruppen)

```bash
python matcher.py
```

Spalte A/B der bearbeiteten Datei werden in Spalte B/C der Vorlage übertragen. Für jede Gruppe wird
der Bereich in Spalte A eingetragen: zuerst per Abgleich mit `Benutzergruppen Zuordnung.xlsx`,
sonst über den Namen (`IT_1...` → IT 1, `IT_0...` → IT-Leitung + Stab), sonst per Nachfrage im
Terminal. Neue Zuordnungen werden in `Benutzergruppen Zuordnung.xlsx` übernommen (mit Sicherung).
Das Ergebnis wird als `filled Auswertung Benutzergruppen ZB IT.xlsx` gespeichert.

## Neuzugänge (comparator.py)

Vergleicht `filled Auswertung Benutzergruppen ZB IT.xlsx` mit `toCompareLastYear.xlsx`
(beide: A = betrifft, B = Gruppe, C = Name, ab Zeile 3).

```bash
python comparator.py
```

Ergebnis ist `Veränderungen.xlsx` mit zwei Reitern: "Neuzugänge" (Kombinationen aus Gruppe und Name,
die es letztes Jahr noch nicht gab) und "Abgänge" (Kombinationen, die es dieses Jahr nicht mehr gibt).

## Zuordnungspaare (zuordnung.py)

Vergleicht Spalte A (IT 1 bis IT 6, "IT-Leitung + Stab", "übergreifend") mit Spalte B ab Zeile 3
und gibt eine Exceltabelle mit einer Spalte pro Bereich aus.

```bash
python zuordnung.py
```

Ohne Argument wird `test.xlsx` aus dem Ordner des Skripts gelesen und das Ergebnis als
`test_Zuordnung.xlsx` gespeichert. Eine andere Datei kann als Argument angegeben werden.

## Abgleich mit matching.xlsx (abgleich.py)

Vergleicht `test_Zuordnung.xlsx` mit `matching.xlsx` (ab Zeile 2: Spalte A = Begriff,
Spalte F = Bereich). Beide Dateien müssen im Ordner des Skripts liegen.

```bash
python abgleich.py
```

Ergebnis ist `Abgleich.xlsx` mit einer Zeile pro Paar und dem Status "in beiden",
"nur in test_Zuordnung" oder "nur in matching". Fehlt ein Paar in einer Liste, zeigt die
Spalte "In der anderen Liste bei", welchem Bereich der Begriff dort zugeordnet ist.

Zusätzlich entsteht `Abweichungen.xlsx` mit nur den auffälligen Begriffen: "Widerspruch"
(unterschiedliche Bereiche in beiden Listen), "fehlt in matching" oder "fehlt in test_Zuordnung".
