"""
Startet die drei Skripte der Benutzergruppen-Auswertung nacheinander:

1. transform.py
2. matcher.py
3. comparator.py

Jedes Skript läuft so, als würde es einzeln mit "python <skript>.py" gestartet.
Bricht ein Skript mit einem Fehler ab, werden die folgenden nicht mehr gestartet.

Aufruf:
    python benutzergruppenAuswertung.py
"""

import subprocess
import sys
from pathlib import Path

ORDNER = Path(__file__).resolve().parent

SKRIPTE = ["transform.py", "matcher.py", "comparator.py"]


def main() -> None:
    for nummer, skript in enumerate(SKRIPTE, start=1):
        print()
        print(f"=== Schritt {nummer}/{len(SKRIPTE)}: {skript} ===", flush=True)
        ergebnis = subprocess.run([sys.executable, str(ORDNER / skript)], cwd=ORDNER)
        if ergebnis.returncode != 0:
            sys.exit(f"Abbruch: {skript} ist mit Fehlercode {ergebnis.returncode} beendet worden.")

    print()
    print("Benutzergruppen-Auswertung abgeschlossen.")


if __name__ == "__main__":
    main()
