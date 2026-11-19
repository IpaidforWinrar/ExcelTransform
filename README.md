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

### aktenplanAuswertung.py

Erstellt `Aktenplan bearbeitet.xlsx` aus dem neuen Aktenplan:

1. `AktenplanNeu.xlsx` (Spalte A bis Q, ab Zeile 3) wird in `empty Aktenplan.xlsx` übertragen, ab Zeile 3
   und zwei Spalten nach rechts (A → C, B → D, ... Q → S). Zeile 1 bekommt die Überschrift
   "Berechtigung im Aktenplan ZB IT - Stand Monat Jahr".
2. Zeilen, deren Schlüssel (Spalte D + E) in `AktenplanAlt.xlsx` fehlt, bekommen in Spalte A "neu" in roter Schrift.
3. Jeder Wert aus Spalte C wird in Spalte A von `Aktenplan Zuordnung.xlsx` gesucht. Bei einem Treffer
   kommt der Wert aus Spalte B der Zuordnung in Spalte B. Steht dort vorher ein anderer Wert, wird er
   überschrieben und in der Logdatei festgehalten.

```bash
python "Aktenplan Auswertung/aktenplanAuswertung.py"
```

Eingabe: `FAVIS DB Aktenplan/Eingabe Aktenplaene`, Ausgabe (inkl. Logdatei `Aktenplan Auswertung.log`):
`FAVIS DB Aktenplan/Ausgabe Aktenplan`. Der Eingabe-Ordner wird nur gelesen: `empty Aktenplan.xlsx` wird in
den Ausgabe-Ordner kopiert und als `Aktenplan bearbeitet.xlsx` bearbeitet. Übertragen werden nur Werte,
Formate und übrige Inhalte der Vorlage bleiben unverändert.

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

## Trafo Freeze-Liste Update

`Trafo Freeze-Liste Update/trafoUpdate.py` überträgt die Werte der `Trafo Freeze-Liste.xlsx` in eine Kopie der
`empty Trafo Freeze-Liste.xlsx`. Gematcht wird im Reiter "SOLL Wertstromzuordnung" über Stellennummer (Spalte D)
und Personalnummer (Spalte F) ab Zeile 3; bei einem Treffer werden die Spalten I bis BJ übernommen.

Ordner liegen neben dem Skript: `update Eingabe` (beide xlsx) und `update Ausgabe` (Ergebnis
`filled Trafo Freeze-Liste.xlsx` und `trafoUpdate_Protokoll.txt` mit nicht gematchten und doppelten Keys).
Schlüssel- und Kopierspalten lassen sich oben im Skript ändern (`KEY_COLUMNS`, `COPY_COLUMNS`, `EXCLUDE_COLUMNS`).
Die Werte werden direkt in die Sheet-XML der kopierten Datei geschrieben (kein openpyxl-Speichern), sodass Datenverbindungen,
Power Query und Formatierung der empty-Liste erhalten bleiben. Benötigt `lxml`.

```bash
python "Trafo Freeze-Liste Update/trafoUpdate.py"
```
