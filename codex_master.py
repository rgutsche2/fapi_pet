#%%
"""Arztbriefe auswerten: gemeinsamer Ablauf für SpA, RA und PsA.

Im bestehenden Ordner fapi_pet speichern. Prompts und Schemata liegen daneben.
In PyCharm diese Datei starten. Einstellungen direkt unten ändern.
"""
import hashlib
import json
import math
import re
import shutil
import subprocess
from datetime import date, datetime
from pathlib import Path

from docx import Document
from docx.text.paragraph import Paragraph
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


# ==================== EINSTELLUNGEN ====================
PROJEKTORDNER = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
INDIKATIONEN = ["SpA", "RA", "PsA"]  # Für nur PsA: ["PsA"]
PATIENTENLISTE = PROJEKTORDNER / "data" / "patients_fapi_pet.xlsx"
BRIEFORDNER = PROJEKTORDNER / "data" / "Arztbriefe_Anonymisiert"
AUSGABEORDNER = PROJEKTORDNER / "output"
PROMPTORDNER = PROJEKTORDNER / "prompts"
SCHEMAORDNER = PROJEKTORDNER / "schemata"
FOLLOW_UP_NUMMER = 1  # Erster FU bei mehreren Scans (z.B. Patienten 13, 23, 45).
NEU_AUSWERTEN = False  # Geänderte Prompts/Schemata werden automatisch neu ausgewertet.
CODEX_TIMEOUT_SEKUNDEN = 600

PHASEN = ("Baseline", "Follow-up")
VERARBEITUNGSVERSION = "2.0-erster-crp-therapiebezogen"
MEDIKAMENT_FELDER = {
    "Wirkstoff": "Text", "Dosis": "Zahl", "Dosiseinheit": "Text",
    "Intervallzahl": "Zahl", "Intervalleinheit": "Text", "Bedarf": "Text",
    "Status": "Text", "Dosierung (Original)": "Text",
}
MEDIKAMENT_AUSWAHL = {
    "Intervalleinheit": ["Stunden", "Tage", "Wochen", "Monate"],
    "Bedarf": ["Ja", "Nein", "Unklar"],
    "Status": ["laufend", "geplant", "pausiert", "beendet", "unklar"],
}


def datum(wert):
    if wert is None or wert == "":
        return None
    if isinstance(wert, datetime):
        return wert.date().isoformat()
    if isinstance(wert, date):
        return wert.isoformat()
    return date.fromisoformat(str(wert).strip()[:10]).isoformat()


def konfiguration_laden(indikation):
    if indikation not in ("SpA", "RA", "PsA"):
        raise ValueError(f"Unbekannte Indikation: {indikation}")
    schema = json.loads((SCHEMAORDNER / f"{indikation}.json").read_text(encoding="utf-8"))
    if schema["indikation"] != indikation:
        raise ValueError("Indikation und Schema passen nicht zusammen.")
    spalten = []

    def hinzufuegen(name, typ, gruppe, breite):
        if not isinstance(name, str) or not name.strip() or typ not in ("Text", "Zahl", "Datum", "Formel", "Liste"):
            raise ValueError("Ungültiger Spaltenname oder Feldtyp im Schema.")
        spalten.append({"name": name, "typ": typ, "gruppe": gruppe, "breite": breite})

    for name, typ, breite in [
        ("Patient", "Zahl", 10), ("Erkrankung", "Text", 12), ("Zeitpunkt", "Text", 15),
        ("Scan-Datum", "Datum", 14), ("Klinisches Datum", "Datum", 16),
        ("Abstand zum Scan (Tage)", "Formel", 18),
    ]:
        hinzufuegen(name, typ, "Identifikation", breite)
    for feld in schema["scores"]:
        hinzufuegen(feld["name"], feld["typ"], "Scores", 18 if feld["typ"] == "Text" else 12)
    for name, typ, breite in [
        ("CRP dokumentiert", "Text", 18),
        ("CRP-Wert", "Zahl", 13), ("CRP-Vergleichszeichen", "Text", 18),
        ("CRP-Einheit", "Text", 14), ("BSG (mm nach 1 h)", "Zahl", 18),
    ]:
        hinzufuegen(name, typ, "Labor", breite)
    hinzufuegen("Laufende Wirkstoffe", "Text", "Therapie", 35)
    for name in schema["klinische_spalten"]:
        hinzufuegen(name, "Text", "Klinische Befunde", 25 if name in schema.get("auswahlwerte", {}) else 55)
    hinzufuegen("Beurteilte Therapie", "Text", "Beurteilung", 35)
    hinzufuegen("Ansprechen", "Text", "Beurteilung", 23)
    hinzufuegen("Krankheitszustand", "Text", "Beurteilung", 25)
    hinzufuegen("Verlauf", "Text", "Beurteilung", 65)
    hinzufuegen("Hinweise", "Liste", "Beurteilung", 65)
    namen = [s["name"] for s in spalten]
    if len(set(namen)) != len(namen):
        raise ValueError("Doppelte Spaltennamen im Schema.")
    felder = [s["name"] for s in spalten if s["name"] not in {
        "Patient", "Erkrankung", "Zeitpunkt", "Scan-Datum", "Abstand zum Scan (Tage)", "Hinweise", "Laufende Wirkstoffe",
    }]
    auswahl = {**schema.get("auswahlwerte", {}),
               "CRP-Vergleichszeichen": ["=", "<", "<=", ">", ">="],
               "CRP dokumentiert": ["Ja", "Nein"],
               "Ansprechen": ["Deutlich", "Partiell", "Keines", "Nicht beurteilbar"],
               "Krankheitszustand": ["Remission", "Gut kontrolliert", "Aktiv", "Unklar"]}
    for name in auswahl:
        if name not in felder:
            raise ValueError(f"Auswahlwerte für unbekanntes Feld: {name}")
    for paar in schema.get("gekoppelte_felder", []):
        if len(paar) != 2 or any(name not in felder for name in paar):
            raise ValueError("Gekoppelte Felder müssen zwei bekannte Feldnamen enthalten.")
    struktur = {"Zeilen": [dict(Zeitpunkt=z, **{f: None for f in felder}, Medikamente=[], Belege={}, Hinweise=[]) for z in PHASEN]}
    allgemein = (PROMPTORDNER / "allgemein.txt").read_text(encoding="utf-8")
    fachlich = (PROMPTORDNER / f"{indikation}.txt").read_text(encoding="utf-8")
    typen = {s["name"]: s["typ"] for s in spalten if s["name"] in felder}
    prompt = allgemein + "\nKRANKHEITSSPEZIFISCHE REGELN:\n" + fachlich
    prompt += "\nFELDTYPEN (fehlend bleibt null):\n" + json.dumps(typen, ensure_ascii=False)
    prompt += "\nERLAUBTE KATEGORIEN:\n" + json.dumps(auswahl, ensure_ascii=False)
    medikament = {**{f: None for f in MEDIKAMENT_FELDER}, "Belege": [], "Hinweise": []}
    prompt += "\nMEDIKAMENTE: je Wirkstoff ein Objekt dieses Aufbaus, sonst []:\n" + json.dumps(medikament, ensure_ascii=False)
    prompt += "\nMEDIKAMENT-FELDTYPEN:\n" + json.dumps(MEDIKAMENT_FELDER, ensure_ascii=False)
    prompt += "\nMEDIKAMENT-KATEGORIEN:\n" + json.dumps(MEDIKAMENT_AUSWAHL, ensure_ascii=False)
    prompt += "\nAUSGABESTRUKTUR:\n" + json.dumps(struktur, ensure_ascii=False, indent=2)
    return {"indikation": indikation, "schema": schema, "spalten": spalten, "felder": felder,
            "zahlenfelder": [s["name"] for s in spalten if s["typ"] == "Zahl" and s["name"] in felder],
            "auswahl": auswahl, "struktur": struktur, "prompt": prompt}


def patienten_laden(indikation):
    buch = load_workbook(PATIENTENLISTE, data_only=True, read_only=True)
    try:
        zeilen = list(buch["Baseline + Follow-up"].values)
    finally:
        buch.close()
    kopf = {str(v).strip(): i for i, v in enumerate(zeilen[0]) if v is not None}
    benoetigt = {"PID", "Studiengruppe", "Scandatum", "Follow-up Daten"}
    if not benoetigt <= set(kopf):
        raise ValueError("Patientenliste: Pflichtspalten fehlen: " + ", ".join(sorted(benoetigt - set(kopf))))
    patienten = []
    for r in zeilen[1:]:
        if r[kopf["PID"]] is None or str(r[kopf["Studiengruppe"]]).strip().casefold() != indikation.casefold():
            continue
        pid = int(r[kopf["PID"]])
        baseline = datum(r[kopf["Scandatum"]])
        fu = [datum(x) for x in str(r[kopf["Follow-up Daten"]] or "").split("|") if x.strip()]
        index = FOLLOW_UP_NUMMER - 1 if len(fu) > 1 else 0
        if baseline is None or not fu or index >= len(fu):
            raise ValueError(f"Patient {pid}: Baseline oder gewählter Follow-up-Scan fehlt.")
        patienten.append({"Patient": pid, "Erkrankung": indikation,
                          "Scans": {"Baseline": baseline, "Follow-up": fu[index]}})
    if len({p["Patient"] for p in patienten}) != len(patienten):
        raise ValueError(f"Doppelte {indikation}-Patientennummer in der Patientenliste.")
    return sorted(patienten, key=lambda p: p["Patient"])


def word_lesen(pfad):
    teile = []
    for block in Document(pfad).iter_inner_content():
        if isinstance(block, Paragraph):
            teile.append(block.text)
        else:
            for zeile in block.rows:
                teile.append(" | ".join(zelle.text for zelle in zeile.cells))
    text = "\n".join(teile)
    if not text.strip():
        raise ValueError(f"Kein lesbarer Text: {pfad.name}")
    return text


def briefe_laden(pid):
    texte, gruppen = {}, {z: [] for z in PHASEN}
    dateien = sorted(p for p in BRIEFORDNER.rglob("*") if p.is_file() and p.suffix.lower() == ".docx")
    for phase, name in (("Baseline", "Baseline"), ("Follow-up", "FollowUp")):
        muster = re.compile(rf"{pid}_{name}(?:_\d+)?(?:\s*\(\d+\))?\.docx", re.I)
        gesehen = set()
        for pfad in dateien:
            if not muster.fullmatch(pfad.name):
                continue
            fingerprint = hashlib.sha256(pfad.read_bytes()).hexdigest()
            if fingerprint in gesehen:
                continue
            gesehen.add(fingerprint)
            quelle = pfad.relative_to(BRIEFORDNER).as_posix()
            texte[quelle] = word_lesen(pfad)
            gruppen[phase].append(quelle)
    if not all(gruppen.values()):
        raise ValueError("Arztbrief fehlt für " + ", ".join(k for k, v in gruppen.items() if not v))
    return texte, gruppen


def prompt_erstellen(config, patient, texte, gruppen):
    prompt = config["prompt"] + "\nPATIENT UND SCAN-DATEN:\n" + json.dumps(patient, ensure_ascii=False)
    crp = {phase: erster_crp_wert(texte, gruppen[phase]) for phase in PHASEN}
    prompt += "\nREGELBASIERT ERKANNTER ERSTER CRP-WERT JE BRIEFGRUPPE:\n" + json.dumps(crp, ensure_ascii=False)
    prompt += "\nErkannte CRP-Werte verbindlich übernehmen; null bedeutet nur: kein einfaches Zahlenformat erkannt.\n"
    for phase in PHASEN:
        for quelle in gruppen[phase]:
            prompt += f"\nBRIEFGRUPPE: {phase}\nDATEI: {quelle}\n{texte[quelle]}\nENDE BRIEF\n"
    # Auch Typen, erlaubte Werte und neue Spalten machen alte Ergebnisse ungültig.
    inhalt = json.dumps({"version": VERARBEITUNGSVERSION, "prompt": prompt, "schema": config["schema"], "spalten": config["spalten"]},
                        ensure_ascii=False, sort_keys=True)
    return prompt, hashlib.sha256(inhalt.encode("utf-8")).hexdigest()


def erster_crp_wert(texte, quellen):
    """Erste klar beschriftete Messung in der übergebenen Brief-/Textreihenfolge.

    Das ist eine Auswahlkonvention, keine chronologische Datierung. Einfache
    CRP-Zahlen werden so unabhängig vom LLM gesichert. Andere Schreibweisen
    verbleiben beim Modell; hsCRP, Scorewerte und Referenzangaben sind kein Ersatz.
    """
    label = r"(?<![\w\-\u2011\u2013])(?:C[\-\u2011\u2013 ]?reaktives?\s+Protein|CRP(?:[ -]Wert)?)(?!\w)"
    muster = re.compile(
        label +
        r"\s*(?:[:|=]\s*)?(?P<zeichen><=|>=|\u2264|\u2265|<|>|=)?\s*"
        r"(?P<wert>\d+(?:[.,]\d+)?)(?![\d.,])"
        r"(?:\s*(?:\|\s*)?(?P<einheit>mg\s*/\s*d?l|g\s*/\s*l))?", re.I,
    )
    for quelle in quellen:
        for name in re.finditer(label, texte[quelle], re.I):
            davor = texte[quelle][max(0, name.start() - 65):name.start()]
            if re.search(r"(?:ASDAS|DAS[ -]?28)[\s(\-\u2011\u2013]*$", davor, re.I):
                continue
            if re.search(r"(?:Referenz(?:bereich|wert)?|Norm(?:al)?wert|Minimum|Maximum|Min\.?|Max\.?)\s*(?:f[üu]r\s*)?$", davor, re.I):
                continue
            treffer = muster.match(texte[quelle], name.start())
            if treffer is None:
                # Ein früheres unbekanntes Format nicht überspringen und dann
                # versehentlich eine spätere Messung verbindlich festschreiben.
                return None
            danach = texte[quelle][treffer.end():treffer.end() + 10]
            if re.match(r"\s*[\-\u2013]\s*\d", danach):  # Ein Zahlenbereich ist kein einzelner Messwert.
                return None
            einheit = treffer.group("einheit")
            if einheit:
                einheit = re.sub(r"\s+", "", einheit).lower()
            zeichen = (treffer.group("zeichen") or "=").replace("\u2264", "<=").replace("\u2265", ">=")
            return {"Wert": float(treffer.group("wert").replace(",", ".")), "Vergleichszeichen": zeichen,
                    "Einheit": einheit, "Datei": quelle, "Zitat": treffer.group(0).strip()}
    return None


def crp_festlegen(zeile, texte, gruppen):
    crp = erster_crp_wert(texte, gruppen[zeile["Zeitpunkt"]])
    if crp is None:
        return
    beleg = [{"Datei": crp["Datei"], "Zitat": crp["Zitat"]}]
    for feld, wert in {"CRP dokumentiert": "Ja", "CRP-Wert": crp["Wert"],
                       "CRP-Vergleichszeichen": crp["Vergleichszeichen"], "CRP-Einheit": crp["Einheit"]}.items():
        zeile[feld] = wert
        if wert is None:
            zeile["Belege"].pop(feld, None)
        else:
            zeile["Belege"][feld] = beleg
    hinweis = "CRP-Auswahl: erster dokumentierter Messwert in der Briefreihenfolge; daraus kein genaues Labordatum abgeleitet."
    if hinweis not in zeile["Hinweise"]:
        zeile["Hinweise"].append(hinweis)


def belege_pruefen(belege, quellen, texte):
    if not isinstance(belege, list):
        raise ValueError("Belege müssen eine Liste sein.")
    for b in belege:
        if not isinstance(b, dict) or set(b) != {"Datei", "Zitat"}:
            raise ValueError("Ein Beleg benötigt Datei und Zitat.")
        if not isinstance(b["Datei"], str) or b["Datei"] not in quellen:
            raise ValueError("Beleg stammt aus der falschen Briefgruppe.")
        if not isinstance(b["Zitat"], str) or not b["Zitat"].strip():
            raise ValueError("Leeres Originalzitat.")
        if " ".join(b["Zitat"].split()) not in " ".join(texte[b["Datei"]].split()):
            raise ValueError(f"Originalzitat nicht gefunden: {b['Zitat']}")


def medikamente_pruefen(medikamente, quellen, texte):
    if not isinstance(medikamente, list):
        raise ValueError("Medikamente muss eine Liste sein.")
    for med in medikamente:
        if not isinstance(med, dict) or set(med) != set(MEDIKAMENT_FELDER) | {"Belege", "Hinweise"}:
            raise ValueError("Unvollständiger oder unerwarteter Medikament-Aufbau.")
        if not med["Wirkstoff"] or med["Status"] is None or not med["Belege"]:
            raise ValueError("Medikament benötigt Wirkstoff, Status und Originalbeleg.")
        if not isinstance(med["Hinweise"], list) or not all(isinstance(x, str) for x in med["Hinweise"]):
            raise ValueError("Medikament-Hinweise müssen eine Textliste sein.")
        for f, typ in MEDIKAMENT_FELDER.items():
            v = med[f]
            if v is None:
                continue
            if typ == "Zahl":
                if isinstance(v, bool) or not isinstance(v, (int, float)) or not math.isfinite(v) or v < 0:
                    raise ValueError(f"Medikament {f}: nichtnegative Zahl erwartet.")
            elif not isinstance(v, str) or not v.strip():
                raise ValueError(f"Medikament {f}: Text erwartet.")
        for f, werte in MEDIKAMENT_AUSWAHL.items():
            if med[f] is not None and med[f] not in werte:
                raise ValueError(f"Ungültige Medikament-Angabe: {f}")
        if (med["Intervallzahl"] is None) != (med["Intervalleinheit"] is None):
            raise ValueError("Intervallzahl und -einheit zusammen angeben.")
        if med["Intervallzahl"] == 0:
            raise ValueError("Ein Medikamentintervall muss größer als null sein.")
        belege_pruefen(med["Belege"], quellen, texte)


def antwort_pruefen(daten, config, texte, gruppen, nur_phase=None):
    if not isinstance(daten, dict) or set(daten) != {"Zeilen"}:
        raise ValueError("Erwartet wird ein JSON-Objekt mit 'Zeilen'.")
    zeilen = daten["Zeilen"]
    if not isinstance(zeilen, list) or len(zeilen) != 2 or not all(isinstance(z, dict) for z in zeilen):
        raise ValueError("Erwartet werden genau zwei Zeilen als JSON-Objekte.")
    if sorted(z.get("Zeitpunkt", "") for z in zeilen if isinstance(z.get("Zeitpunkt"), str)) != sorted(PHASEN):
        raise ValueError("Baseline und Follow-up müssen jeweils einmal vorkommen.")
    gewaehlt = [z for z in zeilen if nur_phase is None or z["Zeitpunkt"] == nur_phase]
    if not gewaehlt:
        raise ValueError("Angeforderte Briefphase fehlt.")
    for z in gewaehlt:
        if set(z) != set(config["felder"]) | {"Zeitpunkt", "Medikamente", "Belege", "Hinweise"}:
            raise ValueError("Unerwartete oder fehlende JSON-Felder.")
        if not isinstance(z["Hinweise"], list) or not all(isinstance(x, str) for x in z["Hinweise"]):
            raise ValueError("Hinweise muss eine Textliste sein.")
        if not isinstance(z["Belege"], dict) or set(z["Belege"]) - set(config["felder"]):
            raise ValueError("Belege muss ein Objekt mit bekannten Feldnamen sein.")
        crp_festlegen(z, texte, gruppen)
        for f in config["felder"]:
            v = z[f]
            if f in config["zahlenfelder"]:
                if v is not None and (isinstance(v, bool) or not isinstance(v, (int, float)) or not math.isfinite(v)):
                    raise ValueError(f"{f}: Zahl oder null erwartet.")
            elif v is not None and (not isinstance(v, str) or not v.strip()):
                raise ValueError(f"{f}: nichtleerer Text oder null erwartet.")
            if f == "Klinisches Datum" and v is not None and datum(v) != v:
                raise ValueError("Klinisches Datum muss YYYY-MM-DD sein.")
            ohne_beleg = ((f == "Verlauf" and v == "Ausgangsbefund")
                          or (f == "CRP dokumentiert" and v == "Nein")
                          or (f == "Ansprechen" and v == "Nicht beurteilbar")
                          or (f == "Krankheitszustand" and v == "Unklar")
                          or v == "nicht erwähnt")
            if v is not None and not ohne_beleg and not z["Belege"].get(f):
                raise ValueError(f"Originalbeleg fehlt für {f}.")
        for f, erlaubt in config["auswahl"].items():
            if z[f] is not None and z[f] not in erlaubt:
                raise ValueError(f"Unbekannter Wert für {f}: {z[f]}")
        for a, b in config["schema"].get("gekoppelte_felder", []):
            if (z[a] is None) != (z[b] is None):
                raise ValueError(f"{a} und {b} müssen zusammen angegeben werden.")
        if z["Zeitpunkt"] == "Baseline" and (z["Ansprechen"] is not None or z["Beurteilte Therapie"] is not None):
            raise ValueError("Ansprechen und beurteilte Therapie werden nur im Follow-up klassifiziert.")
        if z["Zeitpunkt"] == "Follow-up" and z["Ansprechen"] is None:
            raise ValueError("Unklares Ansprechen ausdrücklich als Nicht beurteilbar ausgeben.")
        if z["Ansprechen"] in ("Deutlich", "Partiell", "Keines") and z["Beurteilte Therapie"] is None:
            raise ValueError("Beurteilte Therapie fehlt zum klassifizierten Ansprechen.")
        if z["CRP dokumentiert"] not in ("Ja", "Nein"):
            raise ValueError("CRP dokumentiert muss Ja oder Nein sein.")
        if z["CRP dokumentiert"] == "Nein" and any(z[f] is not None for f in ("CRP-Wert", "CRP-Vergleichszeichen", "CRP-Einheit")):
            raise ValueError("CRP fehlt laut Dokumentationsstatus, aber ein Wert wurde ausgegeben.")
        if z["CRP-Wert"] is not None and z["CRP-Vergleichszeichen"] is None:
            raise ValueError("CRP-Messwert benötigt ein Vergleichszeichen.")
        if z["Krankheitszustand"] is None:
            raise ValueError("Unklaren Krankheitszustand ausdrücklich als Unklar ausgeben.")
        for f in config["schema"].get("status_spalten", []):
            if z[f] is None:
                raise ValueError(f"{f}: nicht erwähnte Befunde ausdrücklich als nicht erwähnt ausgeben.")
        medikamente_pruefen(z["Medikamente"], gruppen[z["Zeitpunkt"]], texte)
        for feld, belege in z["Belege"].items():
            quellen = texte if feld in ("Verlauf", "Ansprechen", "Beurteilte Therapie") else gruppen[z["Zeitpunkt"]]
            belege_pruefen(belege, quellen, texte)
    return sorted(gewaehlt, key=lambda z: z["Zeitpunkt"] != "Baseline")


def codex_antwort(codex, prompt, rohdatei):
    antwort = subprocess.run(
        [codex, "exec", "--skip-git-repo-check", "--sandbox", "read-only", "-"],
        input=prompt, capture_output=True, text=True, encoding="utf-8",
        timeout=CODEX_TIMEOUT_SEKUNDEN, cwd=PROJEKTORDNER,
    )
    if antwort.returncode != 0:
        raise RuntimeError(antwort.stderr[-2000:] or "Codex ohne gültige Antwort beendet.")
    rohdatei.write_text(antwort.stdout, encoding="utf-8")
    return json.loads(antwort.stdout)


def pruefen_mit_korrektur(daten, config, texte, gruppen, codex, prompt, korrekturdatei):
    """Ein Korrekturversuch; danach nur die tatsächlich gültigen Phasen behalten.

    Zitate bleiben streng gegen den Originaltext geprüft. Es werden weder
    ungeprüfte Belege akzeptiert noch Diagnosen oder Antwortkategorien repariert.
    """
    try:
        return antwort_pruefen(daten, config, texte, gruppen), {}, False
    except ValueError as fehler:
        grund = str(fehler)
    print("Antwortprüfung fehlgeschlagen; ein gezielter Korrekturversuch: " + grund, flush=True)
    korrekturprompt = prompt + "\nANTWORT ZUR KORREKTUR:\n" + json.dumps(daten, ensure_ascii=False)
    korrekturprompt += "\nPRÜFFEHLER:\n" + grund
    korrekturprompt += ("\nPrüfe die Antwort erneut an den oben vollständig übergebenen Originalbriefen. "
                        "Korrigiere insbesondere nicht wortgetreue Belegzitate durch tatsächlich im "
                        "Original vorhandene zusammenhängende Zitate. Auch Schreibfehler im Original "
                        "wortgetreu beibehalten. Keine Ähnlichkeit als Beleg erfinden und keine "
                        "Negation, Zahl oder Therapiezuordnung ändern, um die Prüfung zu umgehen. "
                        "Inhaltlich unbelegte Angaben entsprechend den Regeln leer bzw. nicht "
                        "beurteilbar lassen. Gib erneut die vollständigen zwei JSON-Zeilen aus.")
    stoppen = False
    # Die vorherige Antwort behalten, falls der zweite Modellaufruf selbst scheitert.
    kandidat = daten
    try:
        kandidat = codex_antwort(codex, korrekturprompt, korrekturdatei)
        return antwort_pruefen(kandidat, config, texte, gruppen), {}, False
    except (RuntimeError, subprocess.TimeoutExpired, KeyboardInterrupt) as fehler:
        print(f"Korrekturaufruf gestoppt: {type(fehler).__name__}: {fehler}", flush=True)
        stoppen = True
    except (ValueError, TypeError) as fehler:
        print("Auch die korrigierte Antwort ist nicht vollständig gültig: " + str(fehler), flush=True)
    gueltig, fehler_pro_phase = [], {}
    for phase in PHASEN:
        try:
            gueltig.extend(antwort_pruefen(kandidat, config, texte, gruppen, nur_phase=phase))
        except (ValueError, TypeError) as fehler:
            # Falls die Korrektur eine vorher gültige Phase beschädigt hat,
            # darf deren belegte ursprüngliche Fassung erhalten bleiben.
            try:
                gueltig.extend(antwort_pruefen(daten, config, texte, gruppen, nur_phase=phase))
            except (ValueError, TypeError):
                fehler_pro_phase[phase] = f"{type(fehler).__name__}: {fehler}"
    return gueltig, fehler_pro_phase, stoppen


def json_speichern(ziel, daten):
    temp = ziel.with_suffix(".tmp.json")
    temp.write_text(json.dumps(daten, ensure_ascii=False, indent=2), encoding="utf-8")
    temp.replace(ziel)


def excel_speichern(config, zeilen, ordner):
    buch = Workbook()
    blatt = buch.active
    blatt.title = config["indikation"]
    spalten = config["spalten"]
    index = {s["name"]: i for i, s in enumerate(spalten, 1)}
    farben = {"Identifikation": "D9EAF7", "Scores": "E2EFDA", "Labor": "FFF2CC",
              "Therapie": "FCE4D6", "Klinische Befunde": "CBEAF5", "Beurteilung": "D9E2F3"}
    rand = Border(bottom=Side(style="thin", color="D9E2F3"))
    start = 1
    for j, s in enumerate(spalten, 1):
        gruppe = s["gruppe"]
        blatt.column_dimensions[get_column_letter(j)].width = s["breite"]
        c = blatt.cell(2, j, s["name"])
        c.font = Font(name="Calibri", size=11, bold=True)
        c.fill = PatternFill("solid", fgColor=farben[gruppe])
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if j == len(spalten) or spalten[j]["gruppe"] != gruppe:
            if start < j:
                blatt.merge_cells(start_row=1, start_column=start, end_row=1, end_column=j)
            c = blatt.cell(1, start, gruppe)
            c.font = Font(name="Calibri", size=12, bold=True)
            c.fill = PatternFill("solid", fgColor=farben[gruppe])
            c.alignment = Alignment(vertical="center")
            start = j + 1
    for i, z in enumerate(zeilen, 3):
        max_zeilen = 2
        for j, s in enumerate(spalten, 1):
            name = s["name"]
            wert = z.get(name)
            if name == "Laufende Wirkstoffe":
                wert = "; ".join(dict.fromkeys(m["Wirkstoff"] for m in z.get("Medikamente", []) if m["Status"] == "laufend")) or None
            if name == "Hinweise":
                wert = "\n".join(wert or []) or None
            if s["typ"] == "Datum" and wert:
                wert = date.fromisoformat(wert)
            c = blatt.cell(i, j, wert)
            c.font = Font(name="Calibri", size=11)
            c.border = rand
            c.alignment = Alignment(vertical="top", wrap_text=True)
            c.fill = PatternFill("solid", fgColor="F2F6FA" if z["Zeitpunkt"] == "Baseline" else "FFFFFF")
            if s["typ"] == "Datum":
                c.number_format = "yyyy-mm-dd"
            elif s["typ"] in ("Zahl", "Formel"):
                c.number_format = "0" if name in ("Patient", "Abstand zum Scan (Tage)") else "0.###"
            if isinstance(wert, str):
                c.data_type = "s"  # Freitext bleibt auch bei führendem '=' normaler Text.
                breite = max(8, s["breite"] - 3)
                max_zeilen = max(max_zeilen, sum(max(1, math.ceil(len(t) / breite)) for t in wert.split("\n")))
        klinisch = f"{get_column_letter(index['Klinisches Datum'])}{i}"
        scan = f"{get_column_letter(index['Scan-Datum'])}{i}"
        blatt.cell(i, index["Abstand zum Scan (Tage)"], f'=IF({klinisch}="","",{klinisch}-{scan})')
        blatt.row_dimensions[i].height = min(409, max(45, max_zeilen * 15))
    ende = len(zeilen) + 2
    for name, werte in config["auswahl"].items():
        validierung = DataValidation(type="list", formula1='"' + ",".join(werte) + '"', allow_blank=True)
        blatt.add_data_validation(validierung)
        validierung.add(f"{get_column_letter(index[name])}3:{get_column_letter(index[name])}{max(3, ende)}")
    blatt.row_dimensions[1].height = 24
    blatt.row_dimensions[2].height = 46
    blatt.freeze_panes = "D3"
    blatt.auto_filter.ref = f"A2:{get_column_letter(len(spalten))}{ende}"
    therapie = buch.create_sheet("Therapie")
    med_spalten = ["Patient", "Erkrankung", "Zeitpunkt", *MEDIKAMENT_FELDER, "Quelle", "Originalbeleg", "Hinweise"]
    therapie.append(med_spalten)
    for c in therapie[1]:
        c.font = Font(name="Calibri", size=11, bold=True)
        c.fill = PatternFill("solid", fgColor=farben["Therapie"])
        c.alignment = Alignment(wrap_text=True, vertical="center")
    therapie.row_dimensions[1].height = 40
    for z in zeilen:
        for med in z.get("Medikamente", []):
            werte = [z.get("Patient"), z.get("Erkrankung"), z.get("Zeitpunkt")]
            werte += [med[f] for f in MEDIKAMENT_FELDER]
            werte += ["\n".join(dict.fromkeys(b["Datei"] for b in med["Belege"])),
                      "\n".join(dict.fromkeys(b["Zitat"] for b in med["Belege"])),
                      "\n".join(med["Hinweise"]) or None]
            therapie.append(werte)
            for c in therapie[therapie.max_row]:
                c.font = Font(name="Calibri", size=11)
                c.alignment = Alignment(wrap_text=True, vertical="top")
                c.border = rand
                if isinstance(c.value, str):
                    c.data_type = "s"
                elif isinstance(c.value, (float, int)):
                    c.number_format = "0.###"
    for j, name in enumerate(med_spalten, 1):
        therapie.column_dimensions[get_column_letter(j)].width = 65 if name in ("Originalbeleg", "Hinweise") else 40 if name in ("Dosierung (Original)", "Quelle") else 20
    therapie.freeze_panes = "D2"
    therapie.auto_filter.ref = f"A1:{get_column_letter(len(med_spalten))}{max(1, therapie.max_row)}"
    ziel = ordner / f"{config['indikation']}_Auswertung.xlsx"
    temp = ziel.with_suffix(".tmp.xlsx")
    try:
        buch.save(temp)
        temp.replace(ziel)
    finally:
        buch.close()
    return ziel


def auswerten(config, patienten, codex):
    indikation = config["indikation"]
    ordner = AUSGABEORDNER / indikation
    json_ordner = ordner / "json"
    json_ordner.mkdir(parents=True, exist_ok=True)
    alle_zeilen = [{"Patient": p["Patient"], "Erkrankung": indikation, "Zeitpunkt": z,
                    "Scan-Datum": p["Scans"][z], "Verlauf": "Auswertung ausstehend"}
                   for p in patienten for z in PHASEN]
    print(f"\n{indikation}: {len(patienten)} Patienten, {len(alle_zeilen)} Zeilen.", flush=True)
    erfolgreich, stoppen = 0, False
    for nummer, p in enumerate(patienten):
        print(f"[{nummer + 1}/{len(patienten)}] Patient {p['Patient']}", flush=True)
        try:
            texte, gruppen = briefe_laden(p["Patient"])
            prompt, fingerprint = prompt_erstellen(config, p, texte, gruppen)
            ziel = json_ordner / f"{p['Patient']}_{indikation}.json"
            daten = None
            if ziel.exists() and not NEU_AUSWERTEN:
                try:
                    gespeichert = json.loads(ziel.read_text(encoding="utf-8"))
                    if isinstance(gespeichert, dict) and gespeichert.get("fingerprint") == fingerprint:
                        daten = gespeichert["Ergebnis"]
                        print("Unverändertes vorhandenes Ergebnis wird verwendet.")
                except (ValueError, KeyError):
                    print("Vorhandenes JSON nicht lesbar; Patient wird erneut ausgewertet.")
            if daten is None:
                print("Warte auf Codex ...", flush=True)
                daten = codex_antwort(codex, prompt, json_ordner / f"{p['Patient']}_{indikation}_roh.txt")
            zeilen, prueffehler, stoppen = pruefen_mit_korrektur(
                daten, config, texte, gruppen, codex, prompt,
                json_ordner / f"{p['Patient']}_{indikation}_korrektur_roh.txt",
            )
            schluessel = "TeilErgebnis" if prueffehler else "Ergebnis"
            gespeichert = {"Patient": p, "fingerprint": fingerprint, schluessel: {"Zeilen": zeilen}}
            if prueffehler:
                gespeichert["Prueffehler"] = prueffehler
            json_speichern(ziel, gespeichert)
            for z in zeilen:
                offset = PHASEN.index(z["Zeitpunkt"])
                alle_zeilen[2 * nummer + offset].update(z)
                if z["Hinweise"]:
                    print(z["Zeitpunkt"] + " – Hinweise: " + " | ".join(z["Hinweise"]))
            for phase, meldung in prueffehler.items():
                z = alle_zeilen[2 * nummer + PHASEN.index(phase)]
                z["Verlauf"] = "Keine gültige Auswertung dieser Phase: " + meldung
                z["Hinweise"] = ["Andere gültige Phase desselben Patienten bleibt erhalten.", meldung]
                print(phase + " – Prüfung fehlgeschlagen: " + meldung, flush=True)
            if not prueffehler:
                erfolgreich += 1
            else:
                print(f"{len(zeilen)}/2 Phasen gültig; Patient beim nächsten Lauf erneut prüfen.", flush=True)
            print("JSON gespeichert.", flush=True)
        except KeyboardInterrupt:
            print("Unterbrochen; gespeicherte Ergebnisse bleiben erhalten.")
            stoppen = True
        except Exception as fehler:
            meldung = f"{type(fehler).__name__}: {fehler}"
            print("Fehler: " + meldung, flush=True)
            for offset in (0, 1):
                alle_zeilen[2 * nummer + offset]["Verlauf"] = "Keine Auswertung: " + meldung
            if isinstance(fehler, (RuntimeError, subprocess.TimeoutExpired)):
                print("Codex-Lauf gestoppt. Erneut starten, um gespeicherte Ergebnisse weiterzuverwenden.")
                stoppen = True
        finally:
            excel_speichern(config, alle_zeilen, ordner)
        if stoppen:
            break
    print(f"Fertig ausgewertet: {erfolgreich}/{len(patienten)} Patienten.")
    print(f"Excel: {ordner / (indikation + '_Auswertung.xlsx')}")
    print(f"JSON und Originalbelege: {json_ordner}")
    return stoppen


def main():
    if not PATIENTENLISTE.is_file() or not BRIEFORDNER.is_dir():
        raise SystemExit(f"Patientenliste oder Briefordner fehlt. Prüfe {PATIENTENLISTE} und {BRIEFORDNER}.")
    if not INDIKATIONEN or len(set(INDIKATIONEN)) != len(INDIKATIONEN):
        raise SystemExit("INDIKATIONEN darf nicht leer sein oder doppelte Einträge enthalten.")
    if isinstance(FOLLOW_UP_NUMMER, bool) or not isinstance(FOLLOW_UP_NUMMER, int) or FOLLOW_UP_NUMMER < 1:
        raise SystemExit("FOLLOW_UP_NUMMER muss eine positive ganze Zahl sein.")
    # Alle Konfigurationen vor dem ersten Modellaufruf prüfen.
    konfigurationen = [konfiguration_laden(i) for i in INDIKATIONEN]
    listen = {c["indikation"]: patienten_laden(c["indikation"]) for c in konfigurationen}
    codex = shutil.which("codex")
    if codex is None:
        raise SystemExit("Codex nicht gefunden. Nach der Installation PyCharm neu starten.")
    for config in konfigurationen:
        patienten = listen[config["indikation"]]
        if not patienten:
            print(f"Keine {config['indikation']}-Patienten in der Liste; übersprungen.")
            continue
        if auswerten(config, patienten, codex):
            break


if __name__ == "__main__":
    main()
