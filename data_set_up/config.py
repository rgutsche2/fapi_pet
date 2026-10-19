"""
Zentrale Einstellungen für das FAPI-PET-Auswertungsprojekt.

Normalerweise müssen nur INPUT_FILE und ggf. OUTPUT_FILE angepasst werden.
Die Eingabedatei sollte im gleichen Ordner wie main.py liegen.
"""

from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parent

# Eingabedatei im Projektordner
INPUT_FILE = PROJECT_DIR / "Patientenliste_raw.xlsx"

# Ausgabeordner und Ergebnisdatei
OUTPUT_DIR = PROJECT_DIR / "output"
OUTPUT_FILE = OUTPUT_DIR / "Patientenauswahl_FAPI.xlsx"

# Tabellenblatt der Eingabedatei:
# 0 = erstes Tabellenblatt; alternativ kann ein Blattname eingetragen werden.
INPUT_SHEET = 0

# Erwartete Kernspalten
ID_COLUMN = "Spalte1"
NAME_COLUMN = "Name"
BIRTHDATE_COLUMN = "Geburtsdatum"
BASELINE_DATE_COLUMN = "Scandatum"
INDICATION_COLUMN = "Indikation"

# Präfix, anhand dessen Follow-up-Spalten automatisch erkannt werden
FOLLOWUP_COLUMN_PREFIX = "Datum "

# Diese Gruppen werden als separate Tabellen ausgegeben.
TARGET_GROUPS = ("RA", "PsA", "axSpA", "pSpA")

# True: zusätzliche Excel-Datei mit allen Original-Indikationen und Häufigkeiten
EXPORT_INDICATION_OVERVIEW = True

# True: Tabellen mit unklarer Follow-up-Dokumentation separat ausgeben
EXPORT_REVIEW_TABLES = True
