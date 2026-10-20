# FAPI PET – gemeinsame Arztbrief-Auswertung, Version 2

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

Die Excel wird nach jedem Patienten gespeichert. Das Hauptblatt behält genau
zwei Zeilen pro Patient. Im zusätzlichen Blatt `Therapie` steht je Medikament
und Zeitpunkt eine Zeile mit Wirkstoff, Dosis, Einheit, Intervall, Bedarf, Status
und Originalbelegen. Das Hauptblatt zeigt die laufenden Wirkstoffe kompakt an. Die Spaltennamen, Reihenfolge,
Gruppen, Filter und Formeln werden aus dem Schema erzeugt. Eine separate
Excel-Vorlage wird dazu nicht benötigt. Die Ergebnisdatei der jeweiligen
Indikation wird beim Lauf aktualisiert. Nach der Prüfung wird auch das jeweilige
JSON aktualisiert; unvollständige Ergebnisse sind ausdrücklich als TeilErgebnis
gekennzeichnet. Frühere Ergebnisse daher bei Bedarf vor dem Lauf separat sichern.

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

## Neue Spalten und Klassifikation

Die Spalten werden aus den Schemata erzeugt; es gibt keine feste Excel-Spaltenzahl.
Die bisherige Spalte `Ansprechen (Ja/Nein)` wird durch folgende Angaben ersetzt:

| Feld | Inhalt |
| --- | --- |
| Beurteilte Therapie | Wirkstoff/Regime, auf das sich die Ansprechaussage bezieht |
| Ansprechen | Deutlich / Partiell / Keines / Nicht beurteilbar |
| Krankheitszustand | Remission / Gut kontrolliert / Aktiv / Unklar |

Baseline: Ansprechen und beurteilte Therapie bleiben leer. Der Krankheitszustand
wird bei beiden Zeitpunkten dokumentiert. Unklares Follow-up-Ansprechen steht
explizit als `Nicht beurteilbar` in der Tabelle.

Ansprechen beschreibt die klinische Veränderung. Aktive Erkrankung bedeutet
nicht automatisch fehlendes Ansprechen; eine Teilbesserung bei verbleibender
Aktivität ist möglich. Umgekehrt beweist gute aktuelle Kontrolle oder Remission
allein noch keine Veränderung gegenüber Baseline. Damit ist z.B. Remission bei
Ansprechen Nicht beurteilbar möglich. Dies ist eine bewusste Verfeinerung der
bisherigen Ja/Nein-Regel; alte Ja-Angaben werden nicht automatisch umcodiert.

Die beurteilte Therapie ist nicht zwingend die gerade laufende Medikation.
Beispiel: `Von Blinatumomab nicht profitiert; aktuell Abatacept` ergibt Keines für
Blinatumomab, während Abatacept als laufende Therapie im Medikamentblatt steht.
Wenn die Wirkung laut Brief noch abgewartet werden soll: Nicht beurteilbar,
auch bei aktuell dokumentierter Krankheitsaktivität. Klinische Angaben werden
nicht durch PET-/MRT- oder Laborverbesserungen ersetzt.

Die klinischen Statusspalten verwenden `vorhanden`, `verneint`, `unklar` und
`nicht erwähnt`. Nicht erwähnt wird niemals als verneint oder als Null gewertet.
Genaue Regionen, Seiten und klinische Besonderheiten stehen unter `Betroffene
Regionen` und `Klinischer Befund (Details)`; die Hauptmanifestation bleibt kurz.

- SpA: axiale Beschwerden/ISG, periphere Gelenkbeschwerden, Synovitis/Schwellung,
  Enthesitis.
- RA: periphere Gelenkbeschwerden, Synovitis/Schwellung, Morgensteifigkeit.
- PsA: periphere Gelenkbeschwerden, axiale Beschwerden/ISG, Synovitis/Schwellung,
  Enthesitis, Daktylitis, Haut- und Nagelbeteiligung.

Bei PsA bezieht sich Ansprechen auf die muskuloskelettale Erkrankung; eine reine
Hautbesserung genügt nicht. Unterschiedliche Haut- und Gelenkentwicklung bleibt
im Verlauf sichtbar. Scores werden nur bei ausdrücklicher Dokumentation übernommen.

## CRP: immer den ersten Messwert übernehmen

Diese Auswahlkonvention gilt für Baseline und Follow-up:

- Bei `1. Wert; letzter Wert [Minimum; Maximum]` wird immer der erste Wert erfasst.
- Bei mehreren Einzelangaben wird die erste Messwertnennung im Text verwendet.
- Bei mehreren Briefen entscheidet die übergebene, nach Dateipfad sortierte
  Briefreihenfolge. Das ist keine Behauptung über den chronologisch frühesten
  Messwert; kein Messdatum wird aus der Position erfunden.
- Referenzbereiche und Min-/Max-Angaben werden nicht als eigene Messungen gezählt.
- Beispiel `13,0; 6,1 (<5) [Min 6,1; Max 13,0]`: Auswahl 13,0 mg/l.
- Beispiel `0,5; 1,5 [Max 76,3]`: Auswahl 0,5 mg/l, nicht 76,3 mg/l.
- Zwei undatierte CRP-Angaben 36,2 und 25,4: Auswahl 36,2.
- `<5,2 mg/l` bleibt Messwert 5,2 mit Vergleichszeichen `<`; kein exakter Wert 5,2.
- Fehlende Einzeldaten führen nicht zum Ausschluss. Abweichende Daten und die
  Auswahlregel stehen in `Hinweise`.

`CRP dokumentiert` wird separat mit Ja/Nein erfasst. Bei rein qualitativer Angabe
wie `CRP unauffällig` ist Dokumentiert=Ja möglich, während der Zahlenwert leer
bleibt. Ein Referenzbereich oder die Scorebezeichnung DAS28-CRP allein genügt nicht.

Das Master-Skript erkennt einfache, klar beschriftete CRP-Zahlen zusätzlich per
Textregel und legt den ersten Wert vor und nach der LLM-Auswertung verbindlich
fest. Der genaue Originalbeleg wird erhalten. Die Regel überspringt hsCRP,
Scorebezeichnungen und klar als Referenz/Min/Max markierte Angaben. Bei einem
früheren unbekannten Schreibformat wird nicht einfach auf einen späteren Wert
ausgewichen: dann bleibt die Auswahl beim LLM nach derselben Erstwert-Regel.
Komplexe Formate werden daher weiterhin an den Originalbelegen geprüft.

BSG bleibt ein Feld in mm nach einer Stunde gemäß Studienkonvention; bei `11/22`
wird weiterhin der erste Wert übernommen.

## Zitatprüfung und Wiederaufnahme

Jeder inhaltlich ausgefüllte Extraktionswert benötigt einen wortgetreuen Beleg.
Reine Fehlend-Kategorien benötigen keinen künstlichen Beleg. Nach einem
Prüffehler bekommt Codex einmal die vollständigen Originaltexte, die Antwort
und den konkreten Fehler zur Korrektur. Schreibfehler im Original müssen auch
im Zitat wortgetreu bleiben. Unscharfe Zitate werden nicht ungeprüft akzeptiert.
Die Korrekturantwort wird als `<Patient>_<Indikation>_korrektur_roh.txt` gespeichert.

Scheitert die Prüfung weiterhin, werden Baseline und Follow-up getrennt geprüft.
Eine gültige Phase bleibt erhalten; nur die tatsächlich fehlerhafte Phase erhält
in Excel einen Prüfhinweis. Das JSON enthält dann `TeilErgebnis` und `Prueffehler`
anstelle eines vollständigen `Ergebnis`. Solche Patienten werden beim nächsten
Lauf erneut angefordert. Die Abschlusszahl zählt nur vollständig gültige Patienten.

Alle geänderten Prompts, Schemata und die neue Verarbeitungsversion lösen eine
Neuauswertung bisheriger Ergebnisse aus. `NEU_AUSWERTEN=False` kann stehen bleiben.
Danach werden unveränderte vollständige Ergebnisse wiederverwendet. Die bisherigen
Ergebnis-Exceldateien werden beim Lauf aktualisiert; manuelle Korrekturen sollten
vorher in einer separaten Kopie aufbewahrt werden.

Technisch mit simulierten Codex-Antworten geprüft; die CRP-Textregel zusätzlich
an den verfügbaren Originalbeispielen. Diese Tests belegen keine allgemeine
klinische Extraktionsgüte des LLM. Originaltexte und Belege bleiben die Grundlage
der inhaltlichen Prüfung.
