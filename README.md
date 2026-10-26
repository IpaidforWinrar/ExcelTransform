# ExcelTransform

Python-Skripte (openpyxl) zur Auswertung von Benutzergruppen und Sicherheitsfreigaben.

```
Benutzergruppen Auswertung/          benutzergruppenAuswertung.py, transform.py, matcher.py, comparator.py
Sicherheitsfreigabe Auswertung/      sicherheitsfreigabe.py
FAVIS DB Benutzergruppen/
├── Eingabe Benutzergruppen/         Auswertung Benutzergruppen.xlsx, empty Auswertung Benutzergruppen ZB IT.xlsx,
│                                    Benutzergruppen Zuordnung.xlsx, toCompareLastYear.xlsx
└── Ausgabe Benutzergruppen/         Ergebnisse (wird angelegt)
FAVIS DB Sicherheitsfreigabe/
├── Eingabe Sicherheitsfreigabe/     Auswertung Sicherheitsfreigabe.xlsx, empty Auswertung Sicherheitsfreigabe.xlsx
└── Ausgabe Sicherheitsfreigabe/     Ergebnisse und Logdatei (wird angelegt)
```

Exceldateien und Logdateien werden nicht eingecheckt.

## Aktenplan Auswertung

`Aktenplan Auswertung/aktenplan.py` liest `AktenplanAlt.xlsx` (ab Zeile 3: Spalte B = Bereich,
Spalte C = Aktenplan-Eintrag) und prüft, ob jeder Eintrag aus Spalte C nur einem Bereich zugeordnet
ist ("unique") oder mehreren ("mehrfach").

```bash
python "Aktenplan Auswertung/aktenplan.py"
```

Ergebnis ist `Aktenplan Zuordnung.xlsx` mit einer Zeile pro Eintrag (Status, Bereiche, Anzahl Bereiche,
Anzahl Zeilen, Zeilennummern). Mehrfach zugeordnete Einträge stehen oben.

Zusätzlich wird auf Widersprüche geprüft: Nennt der Eintrag in Spalte C selbst einen Bereich
(z. B. "IT 1 Server"), ist aber eindeutig einem anderen Bereich zugeordnet (z. B. IT 7), steht in der
Spalte "Widerspruch" ein "ja" mit Erklärung. Ein Bereich zählt nur als eigenes Wort: "IT 1", "it1",
"IT_1", "IT-1" und "IT-Leitung" werden erkannt, "IT 10", "WIT 1" und "IT.01.01.01" nicht.

## Installation

```bash
pip install -r requirements.txt
```

# Benutzergruppen Auswertung

Die komplette Auswertung läuft mit einem Befehl:

```bash
python "Benutzergruppen Auswertung/benutzergruppenAuswertung.py"
```

Das führt nacheinander `transform.py` → `matcher.py` → `comparator.py` aus. Bricht ein Schritt ab,
z. B. weil eine Datei fehlt, werden die folgenden Schritte nicht ausgeführt. Die einzelnen Skripte
lassen sich weiterhin auch einzeln starten.

## transform.py

Bearbeitet `Auswertung Benutzergruppen.xlsx`:

1. Löscht die ersten 3 Zeilen.
2. Merged Spalte B in Spalte C: Ist die Zelle in C leer, wird der Wert aus B übernommen.
3. Löscht Spalte B, dadurch wird Spalte C zur neuen Spalte B.

Ergebnis: `Ausgabe Benutzergruppen/Auswertung Benutzergruppen_bearbeitet.xlsx`

## matcher.py

Überträgt Spalte A/B von `Auswertung Benutzergruppen_bearbeitet.xlsx` in Spalte B/C der Vorlage
`empty Auswertung Benutzergruppen ZB IT.xlsx` (A = betrifft, B = Gruppe, C = Name, ab Zeile 3).
Für jede Gruppe wird der Bereich in Spalte A eingetragen: zuerst per Abgleich mit
`Benutzergruppen Zuordnung.xlsx`, sonst über den Namen (`IT_1...` → IT 1, `IT_0...` → IT-Leitung + Stab),
sonst per Nachfrage im Terminal. Neue Zuordnungen werden in `Benutzergruppen Zuordnung.xlsx`
übernommen (mit Sicherung).

Ergebnis: `Ausgabe Benutzergruppen/filled Auswertung Benutzergruppen ZB IT.xlsx`

## comparator.py

Vergleicht `filled Auswertung Benutzergruppen ZB IT.xlsx` mit `toCompareLastYear.xlsx`
(beide: A = betrifft, B = Gruppe, C = Name, ab Zeile 3).

Ergebnis: `Ausgabe Benutzergruppen/Veränderungen.xlsx` mit zwei Reitern: "Neuzugänge" (Kombinationen
aus Gruppe und Name, die es letztes Jahr noch nicht gab) und "Abgänge" (Kombinationen, die es dieses
Jahr nicht mehr gibt).

# Sicherheitsfreigabe Auswertung

## sicherheitsfreigabe.py

Benötigt:

- `Auswertung Sicherheitsfreigabe.xlsx` (Spalten A bis H, Zeile 1 = Überschrift, Gruppen in Spalte H mit "," getrennt)
- `empty Auswertung Sicherheitsfreigabe.xlsx` (Vorlage, Überschriften in Zeile 2, Werte ab Zeile 3)
- `Benutzergruppen Zuordnung.xlsx` aus `FAVIS DB Benutzergruppen/Eingabe Benutzergruppen`

```bash
python "Sicherheitsfreigabe Auswertung/sicherheitsfreigabe.py"
```

Ordnet jede Zeile anhand der Gruppen in Spalte H den Bereichen aus `Benutzergruppen Zuordnung.xlsx`
zu (oder über `IT_1` bis `IT_7`, `IT_0` → IT-Leitung + Stab). Gruppen ohne Treffer werden ignoriert,
die Zuordnungsdatei wird nicht verändert. Mehrere Bereiche werden in Spalte J "Doppelzuordnung" notiert.

Ergebnis sind 10 Dateien `Auswertung Sicherheitsfreigabe und Gruppenzuordnung ZB IT - X.xlsx`
(X = IT 1 bis IT 7, IT-L_Stab, IT-übergreifend, IT - gesamt) in `Ausgabe Sicherheitsfreigabe`.
Die Gesamtliste hat zusätzlich Spalte K "Zuordnung". Dort liegt auch `Auswertung Sicherheitsfreigabe.log`
mit allem, was im Terminal ausgegeben wurde.
