"""
Führt die komplette Benutzergruppen-Auswertung in einem Durchgang aus:

1. transform.py   "Auswertung Benutzergruppen.xlsx" -> "Auswertung Benutzergruppen_bearbeitet.xlsx"
2. matcher.py     befüllt die Vorlage und ordnet die Bereiche zu
                  -> "filled Auswertung Benutzergruppen ZB IT.xlsx"
3. comparator.py  Vergleich mit dem Vorjahr -> "Veränderungen.xlsx"

Bricht ein Schritt ab (z. B. weil eine Datei fehlt), werden die folgenden Schritte nicht ausgeführt.

Aufruf:
    python benutzergruppenAuswertung.py
"""

import sys

import comparator
import matcher
import transform

SCHRITTE = [
    ("Transformation", transform.main),
    ("Zuordnung", matcher.main),
    ("Vergleich mit Vorjahr", comparator.main),
]


def main() -> None:
    # transform.main() wertet sys.argv aus, daher keine Argumente weiterreichen
    sys.argv = sys.argv[:1]

    for nummer, (titel, schritt) in enumerate(SCHRITTE, start=1):
        print()
        print(f"=== Schritt {nummer}/{len(SCHRITTE)}: {titel} ===")
        try:
            schritt()
        except SystemExit as exc:
            if exc.code not in (None, 0):
                if isinstance(exc.code, str):
                    print(exc.code, file=sys.stderr)
                sys.exit(f"Abbruch in Schritt {nummer} ({titel}).")

    print()
    print("Benutzergruppen-Auswertung abgeschlossen.")


if __name__ == "__main__":
    main()
