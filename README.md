# FAPI PET Patient Database Processing

Autor: **Robin Gutsche**  
Institution: Nuklearmedizin, FAU Erlangen-Nürnberg

## Zweck

Das Projekt liest die Excel-Datei `Patientenliste_raw.xlsx` ein und erstellt eine
formatierte Excel-Ausgabe mit getrennten Tabellen für:

- Rheumatoide Arthritis (`RA`)
- Psoriasis-Arthritis (`PsA`)
- axiale Spondyloarthritis (`axSpA`)
- periphere Spondyloarthritis (`pSpA`)

Zusätzlich werden Follow-up-Untersuchungen erkannt. Da in der ursprünglichen
Patientenliste sowohl Datumswerte als auch Freitextnotizen in den Follow-up-Spalten
stehen können, unterscheidet das Skript zwischen:

- sicherem Follow-up mit Datum,
- Notiz zu einem Follow-up,
- unklarem Follow-up-Hinweis mit fehlendem Datum.

Unklare Fälle werden nicht automatisch als sicherer zusätzlicher Scan gewertet,
sondern zur manuellen Prüfung markiert.

## Projektstruktur

```text
FAPI_DataProcessing/
├── main.py
├── config.py
├── indication_mapping.py
├── followup_detection.py
├── processing.py
├── excel_output.py
├── requirements.txt
├── Patientenliste_raw.xlsx
└── output/
```

## Installation in PyCharm

1. Den Projektordner in PyCharm öffnen.
2. Die Datei `Patientenliste_raw.xlsx` in den Projektordner kopieren.
3. In PyCharm unten das Terminal öffnen.
4. Abhängigkeiten installieren:

```bash
pip install -r requirements.txt
```

5. `data_set_up/main.py` ausführen.

## Ausgabe

Die Ergebnisdatei wird hier erstellt:

```text
output/Patientenauswahl_FAPI.xlsx
```

Sie enthält unter anderem:

- `PsA_Alle`
- `PsA_Followups`
- `PsA_Followup_Long`
- `PsA_Pruefung`
- entsprechende Tabellen für `SpA` und `axSpA`
- `Nicht_zugeordnet`
- `Indikationen_Uebersicht`

## Neue Indikationsbezeichnungen ergänzen

Noch nicht erkannte Schreibweisen erscheinen in:

```text
Indikationen_Uebersicht
```

Die Regeln werden in `data_set_up/indication_mapping.py` ergänzt.

Beispiel:

```python
EXACT_MAPPING = {
    "neue schreibweise": "PsA",
}
```

oder als Regex-Regel in `PATTERN_RULES`.

## Wichtiger methodischer Hinweis

Die automatische Follow-up-Erkennung basiert auf der vorhandenen Tabellenstruktur.
Bei Freitext ohne eindeutig zugehöriges Datum ist nicht sicher entscheidbar, ob es
sich um einen eigenen Scan oder lediglich um eine Notiz handelt. Diese Fälle werden
deshalb unter `Manuelle Follow-up-Prüfung` markiert.

## Weitergabe

Der gesamte Projektordner kann weitergegeben werden. Auf einem anderen Rechner muss
nur Python mit den Paketen aus `requirements.txt` installiert sein. Absolute Pfade
werden nicht verwendet.
