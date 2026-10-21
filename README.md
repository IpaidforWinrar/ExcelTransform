# ExcelTransform

Python-Skripte (openpyxl) zur Auswertung von Benutzergruppen und Sicherheitsfreigaben.

```
Benutzergruppen Auswertung/      transform.py, matcher.py, comparator.py, zuordnung.py, abgleich.py
Sicherheitsfreigabe Auswertung/  sicherheitsfreigabe.py
```

Die Exceldateien gehören jeweils in den Ordner des Skripts. Exceldateien werden nicht eingecheckt.

## Installation

```bash
pip install -r requirements.txt
```

# Benutzergruppen Auswertung

Ablauf: `transform.py` → `matcher.py` → `comparator.py`

## transform.py

Bearbeitet `Auswertung Benutzergruppen.xlsx`:

1. Löscht die ersten 3 Zeilen.
2. Merged Spalte B in Spalte C: Ist die Zelle in C leer, wird der Wert aus B übernommen.
3. Löscht Spalte B, dadurch wird Spalte C zur neuen Spalte B.

```bash
python "Benutzergruppen Auswertung/transform.py"
```

Das Ergebnis wird als `Auswertung Benutzergruppen_bearbeitet.xlsx` gespeichert. Optional können
Eingabe- und Ausgabedatei als Argumente angegeben werden.

## matcher.py

Benötigt:

- `Auswertung Benutzergruppen_bearbeitet.xlsx` (Ergebnis von `transform.py`)
- `empty Auswertung Benutzergruppen ZB IT.xlsx` (Vorlage: A = betrifft, B = Gruppe, C = Name, ab Zeile 3)
- `Benutzergruppen Zuordnung.xlsx` (Zeile 1 = Bereiche, darunter die Gruppen)

```bash
python "Benutzergruppen Auswertung/matcher.py"
```

Spalte A/B der bearbeiteten Datei werden in Spalte B/C der Vorlage übertragen. Für jede Gruppe wird
der Bereich in Spalte A eingetragen: zuerst per Abgleich mit `Benutzergruppen Zuordnung.xlsx`,
sonst über den Namen (`IT_1...` → IT 1, `IT_0...` → IT-Leitung + Stab), sonst per Nachfrage im
Terminal. Neue Zuordnungen werden in `Benutzergruppen Zuordnung.xlsx` übernommen (mit Sicherung).
Das Ergebnis wird als `filled Auswertung Benutzergruppen ZB IT.xlsx` gespeichert.

## comparator.py

Vergleicht `filled Auswertung Benutzergruppen ZB IT.xlsx` mit `toCompareLastYear.xlsx`
(beide: A = betrifft, B = Gruppe, C = Name, ab Zeile 3).

```bash
python "Benutzergruppen Auswertung/comparator.py"
```

Ergebnis ist `Veränderungen.xlsx` mit zwei Reitern: "Neuzugänge" (Kombinationen aus Gruppe und Name,
die es letztes Jahr noch nicht gab) und "Abgänge" (Kombinationen, die es dieses Jahr nicht mehr gibt).

## zuordnung.py

Vergleicht Spalte A (IT 1 bis IT 6, "IT-Leitung + Stab", "übergreifend") mit Spalte B ab Zeile 3
und gibt eine Exceltabelle mit einer Spalte pro Bereich aus.

```bash
python "Benutzergruppen Auswertung/zuordnung.py"
```

Ohne Argument wird `test.xlsx` aus dem Ordner des Skripts gelesen und das Ergebnis als
`test_Zuordnung.xlsx` gespeichert. Eine andere Datei kann als Argument angegeben werden.

## abgleich.py

Vergleicht `test_Zuordnung.xlsx` mit `matching.xlsx` (ab Zeile 2: Spalte A = Begriff,
Spalte F = Bereich).

```bash
python "Benutzergruppen Auswertung/abgleich.py"
```

Ergebnis ist `Abgleich.xlsx` mit einer Zeile pro Paar und dem Status "in beiden",
"nur in test_Zuordnung" oder "nur in matching". Fehlt ein Paar in einer Liste, zeigt die
Spalte "In der anderen Liste bei", welchem Bereich der Begriff dort zugeordnet ist.

Zusätzlich entsteht `Abweichungen.xlsx` mit nur den auffälligen Begriffen: "Widerspruch"
(unterschiedliche Bereiche in beiden Listen), "fehlt in matching" oder "fehlt in test_Zuordnung".

# Sicherheitsfreigabe Auswertung

## sicherheitsfreigabe.py

```
Sicherheitsfreigabe Auswertung/
├── sicherheitsfreigabe.py
├── Eingabe Auswertung/   die 3 benötigten Dateien
└── Ausgabe Auswertung/   Ergebnisse und Logdatei (wird angelegt)
```

Benötigt im Ordner `Eingabe Auswertung`:

- `Auswertung Sicherheitsfreigabe.xlsx` (Spalten A bis H, Zeile 1 = Überschrift, Gruppen in Spalte H mit "," getrennt)
- `empty Auswertung Sicherheitsfreigabe.xlsx` (Vorlage, Überschriften in Zeile 2, Werte ab Zeile 3)
- `Benutzergruppen Zuordnung.xlsx` (wird ersatzweise auch im Ordner "Benutzergruppen Auswertung" gefunden)

```bash
python "Sicherheitsfreigabe Auswertung/sicherheitsfreigabe.py"
```

Ordnet jede Zeile anhand der Gruppen in Spalte H den Bereichen aus `Benutzergruppen Zuordnung.xlsx`
zu (oder über `IT_1` bis `IT_7`, `IT_0` → IT-Leitung + Stab). Gruppen ohne Treffer werden ignoriert,
die Zuordnungsdatei wird nicht verändert. Mehrere Bereiche werden in Spalte J "Doppelzuordnung" notiert.

Ergebnis sind 10 Dateien `Auswertung Sicherheitsfreigabe und Gruppenzuordnung ZB IT - X.xlsx`
(X = IT 1 bis IT 7, IT-L_Stab, IT-übergreifend, IT - gesamt) im Unterordner `Ausgabe Auswertung`.
Die Gesamtliste hat zusätzlich Spalte K "Zuordnung". Dort liegt auch `Auswertung Sicherheitsfreigabe.log`
mit allem, was im Terminal ausgegeben wurde.
