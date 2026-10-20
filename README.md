# FAPI PET – gemeinsame Arztbrief-Auswertung

Die Dateien aus diesem Ordner in den bestehenden Projektordner `fapi_pet` kopieren.
Die vorhandenen Ordner `data` und `output` bleiben dort. Anschließend in PyCharm
`codex_master.py` starten. Die bisherigen drei `codex_test`-Skripte werden für
diesen Ablauf nicht mehr benötigt und können als alte Versionen aufbewahrt werden.

Einmalig im PyCharm-Terminal:

```bash
python -m pip install -r requirements.txt
```

Der bereits verwendete Codex-CLI muss installiert und angemeldet sein.

## Einstellungen

Oben in `codex_master.py` stehen die Einstellungen:

```python
INDIKATIONEN = ["SpA", "RA", "PsA"]  # Alle drei nacheinander
# INDIKATIONEN = ["PsA"]           # Alternativ nur PsA
FOLLOW_UP_NUMMER = 1                # Bei mehreren FU-Scans den ersten nehmen
NEU_AUSWERTEN = False               # Unveränderte fertige Ergebnisse wiederverwenden
```

Die aktuelle Patientenliste enthält 16 SpA-, 15 RA- und 17 PsA-Patienten.
Pro Patient entstehen genau zwei Zeilen: Baseline und Follow-up. Das Skript
filtert anhand der Studiengruppe; die Patientenzahlen sind nicht fest codiert.

## Eingaben und Ausgabe

- Patientenliste: `data/patients_fapi_pet.xlsx`, Blatt `Baseline + Follow-up`.
- Arztbriefe: `data/Arztbriefe_Anonymisiert`, einschließlich Unterordnern.
- Dateinamen: zum Beispiel `17_Baseline.docx` und `17_FollowUp.docx`.
  Auch zusätzliche Briefe wie `20_Baseline_2.docx` und Dateikopien mit `(1)`
  werden erkannt; bytegleiche Kopien werden pro Phase nur einmal gelesen.
- Excel: `output/SpA/SpA_Auswertung.xlsx`, entsprechend für RA und PsA.
- JSON mit Originalbelegen: `output/<Indikation>/json/<Patient>_<Indikation>.json`.
- Ungeprüfte Modellantworten: daneben als `<Patient>_<Indikation>_roh.txt`.

Die Excel wird nach jedem Patienten gespeichert. Die Spaltennamen, Reihenfolge,
Gruppen, Filter und Formeln werden aus dem Schema erzeugt. Eine separate
Excel-Vorlage wird dazu nicht benötigt. Die Ergebnisdatei der jeweiligen
Indikation wird beim Lauf aktualisiert; vorhandene JSON-Dateien bleiben bis zu
einem erfolgreichen neuen Ergebnis erhalten.

## Prompts und Spalten ändern

- `prompts/allgemein.txt`: gemeinsame Regeln zu Daten, Labor, Therapie,
  Ansprechen und Originalbelegen.
- `prompts/SpA.txt`, `RA.txt`, `PsA.txt`: krankheitsspezifische klinische Regeln.
- `schemata/SpA.json`, `RA.json`, `PsA.json`: krankheitsspezifische Scorefelder,
  klinische Spalten und erlaubte Auswahlwerte.

Die gemeinsame Verarbeitung steht ausschließlich in `codex_master.py`.
Feldnamen im jeweiligen Prompt müssen mit den Namen im Schema übereinstimmen.
Die konkrete JSON-Ausgabestruktur und Feldtypen fügt das Master-Skript selbst an.
Ein neues klinisches Textfeld wird in `klinische_spalten` ergänzt und im
krankheitsspezifischen Prompt beschrieben. Der Excel-Export passt sich an.

Bei geänderten Prompts, Schemata, Brieftexten oder Scan-Daten wird der betreffende
Patient automatisch neu ausgewertet. Ein unveränderter Lauf verwendet fertige
Ergebnisse wieder. Alte Ergebnisse aus den drei Einzel-Skripten werden wegen der
neuen Promptstruktur einmal neu angefordert. Bei einem Codex-Verbindungs- oder
Anmeldefehler stoppt der gesamte Lauf, einschließlich weiterer Indikationen.
Nach Behebung des Fehlers `codex_master.py` erneut starten.

## PsA-Symptomatik

PsA erhält getrennte Spalten für:

1. Klinische Hauptmanifestation.
2. Periphere Gelenkbeschwerden einschließlich Untersuchungsbefund und weiterer Gelenke.
3. Axiale Beschwerden / ISG.
4. Enthesitis mit Lokalisation und Seite.
5. Daktylitis mit betroffenem Finger bzw. betroffener Zehe.
6. Hautbeteiligung.
7. Nagelbeteiligung.

Verlauf, Ansprechen und Hinweise folgen anschließend. Nicht erwähnte Symptome
bleiben leer. Ausdrücklich verneinte Symptome werden als Text erhalten. Die
PsA-Scorefelder sind DAPSA, cDAPSA, TJC, SJC, LEI, PASI und HAQ; Werte werden nur
übernommen, wenn sie im Brief dokumentiert sind.

## Vereinbarte Regeln

- Ein abweichendes Labor- oder Untersuchungsdatum allein führt nicht zum
  Ausschluss eines dokumentierten Werts aus der passenden Briefgruppe.
- Das klinische Datum ist ein Referenztermin. Abweichende Einzeldaten und
  Unsicherheiten stehen in der sichtbaren Excel-Spalte `Hinweise` und im JSON.
- `Ansprechen (Ja/Nein)`: Im Follow-up zählen dokumentierte klinische Besserung,
  Teilbesserung oder ausdrücklich gute/ausreichende klinische Kontrolle unter
  Therapie als `Ja`. Ein ausdrücklicher Baseline-Vergleich ist nicht erforderlich.
  Eindeutig fehlender Nutzen zählt als `Nein`; unklare Fälle bleiben leer.
  Baseline bleibt in dieser Spalte leer. Bei PsA zählt das muskuloskelettale
  Ansprechen; eine reine Hautbesserung reicht nicht aus.
- Keine Ansprechen-Klassifikation allein aus Labor oder PET/MRT, keine
  selbst berechneten Scores. Jedes ausgefüllte Extraktionsfeld benötigt Belege.
- BSG: ein Feld, nach der vereinbarten Studienkonvention in mm nach einer Stunde;
  bei zwei Werten wie `11/22` wird der erste übernommen.

Technisch mit simulierten Codex-Antworten geprüft. Die echte Extraktionsqualität
wird durch Sichtprüfung der Arztbriefe und der gespeicherten Originalbelege
beurteilt; sie wurde durch die technischen Tests nicht nachgewiesen.
