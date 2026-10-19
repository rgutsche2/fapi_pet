"""
Zuordnung der Original-Indikationen zu vier rheumatologischen Zielgruppen:

- RA: Rheumatoide Arthritis
- PsA: Psoriasis-Arthritis
- axSpA: axiale Spondyloarthritis
- pSpA: periphere Spondyloarthritis

Wissenschaftlich wichtige Regel:
Eindeutige Diagnosen werden direkt zugeordnet. Verdachtsdiagnosen, Prä-Stadien,
Overlap-/DD-Angaben und unspezifische "SpA"-Angaben erhalten zusätzlich einen
Zuordnungsstatus. Dadurch können sie später separat geprüft oder aus einer
streng definierten Kohorte ausgeschlossen werden.
"""

from __future__ import annotations

import re
import unicodedata
from typing import Optional


def normalize_text(value: object) -> str:
    """Vereinheitlicht Schreibweisen für robuste Vergleiche."""
    if value is None:
        return ""

    text = str(value).strip().lower()
    if not text or text == "nan":
        return ""

    text = unicodedata.normalize("NFKC", text)
    text = text.replace("–", "-").replace("—", "-").replace("_", " ")
    text = re.sub(r"\s+", " ", text)
    return text


# ---------------------------------------------------------------------------
# 1) Eindeutige / bestätigte Diagnosen
# ---------------------------------------------------------------------------

CONFIRMED_MAPPING = {
    # Rheumatoide Arthritis
    "ra": "RA",
    "rheumatoide arthritis": "RA",
    "ra mit pulmonaler beteiligung": "RA",
    "ra und mukoviszidose": "RA",
    "ra unter mtx": "RA",
    "ra, multi drug resistance": "RA",
    "ra, pulmonale beteillung": "RA",
    "ra/lungenfibrose": "RA",
    "rheumatoide arthritis mit v.a. ild": "RA",
    "seronegative ra": "RA",
    "seropostive ra": "RA",
    "therapierefraktäre ra": "RA",

    # Psoriasis-Arthritis
    "psa": "PsA",
    "polaris, psa": "PsA",
    "psa; hautläsionen": "PsA",
    "psoriasisarthritis": "PsA",

    # Axiale Spondyloarthritis
    "axiale spondylarthropathie": "axSpA",
    "ankylosierende spondylitis": "axSpA",
    "spondylitis ankylosans": "axSpA",
    "bechterew": "axSpA",
    "m.bechterew": "axSpA",
    "morbus bechterew": "axSpA",

    # Periphere Spondyloarthritis
    "periphere spa": "pSpA",
    "ced-assoziierte spa": "pSpA",
    "spa (morbus crohn)": "pSpA",
    "spa mit augenbeteilung": "pSpA",
    "spa und hidradenitis supportiva": "pSpA",
    "spondylarthropathie/tendinopathie": "pSpA",
}


# ---------------------------------------------------------------------------
# 2) Verdacht / Prä-Stadium
#    Wird einer Zielgruppe zugeordnet, aber ausdrücklich markiert.
# ---------------------------------------------------------------------------

SUSPECTED_MAPPING = {
    # RA
    "pre-ra": "RA",
    "pre-ra": "RA",
    "prä-ra": "RA",
    "prä-ra (jakibra)": "RA",
    "prä-ra studie": "RA",
    "v.a. rheumatoide arthritis": "RA",

    # PsA
    "pra-psa": "PsA",
    "prä-psa": "PsA",
    "prä-psa (pso)": "PsA",
    "prä-psa": "PsA",
    "pso, v.a. psa": "PsA",
    "v.a. psa": "PsA",
    "v.a.psa": "PsA",

    # axSpA
    "v.a. ankylosierende spondylitis": "axSpA",
    "v.a. bechterew": "axSpA",

    # pSpA / nicht näher spezifizierte SpA
    "v.a. spa": "pSpA",
}


# ---------------------------------------------------------------------------
# 3) Overlap / Differentialdiagnose / nicht eindeutig
#    Keine automatische Aufnahme in eine reine Zielkohorte.
# ---------------------------------------------------------------------------

AMBIGUOUS_MAPPING = {
    "overlap ra und spa": "RA/SpA",
    "ra dd:psa": "RA/PsA",
    "ra/mischkollagenose": "RA/Overlap",
    "ra/psa overlap": "RA/PsA",
    "ra/periphere spa": "RA/pSpA",
    "seronegative ra/spa": "RA/SpA",
    "spa dd: seronegative ra": "RA/SpA",
    "spa/v.a. psa": "SpA/PsA",
    "v.a.seronegative ra dd: periphere spa": "RA/pSpA",
    "periphere spa dd: seronegative ra": "RA/pSpA",
    "vaskulitis/spa": "SpA/Overlap",
    "psc/spa": "SpA/Overlap",
}


# ---------------------------------------------------------------------------
# 4) Unspezifische Spondyloarthritis-Bezeichnungen
#    "SpA" allein erlaubt keine sichere Trennung axial vs. peripher.
# ---------------------------------------------------------------------------

UNSPECIFIED_SPA = {
    "spa",
    "spa polaris",
    "spa externer patient",
    "spondylarthritis",
    "spondylarthropathie",
    "seronegative spondylarthritis",
}


def classify_indication(value: object) -> Optional[str]:
    """
    Gibt die Zielgruppe zurück.

    Für eindeutige und Verdachtsdiagnosen:
        RA, PsA, axSpA oder pSpA

    Für Overlap/DD und unspezifische SpA:
        None, damit diese Fälle nicht unbemerkt in eine reine Kohorte gelangen.
    """
    text = normalize_text(value)

    if not text:
        return None

    if text in CONFIRMED_MAPPING:
        return CONFIRMED_MAPPING[text]

    if text in SUSPECTED_MAPPING:
        return SUSPECTED_MAPPING[text]

    return None


def classify_indication_status(value: object) -> str:
    """Beschreibt die Sicherheit und Art der Zuordnung."""
    text = normalize_text(value)

    if not text:
        return "keine Indikation"

    if text in CONFIRMED_MAPPING:
        return "eindeutig/bestätigt"

    if text in SUSPECTED_MAPPING:
        return "Verdacht/Prä-Stadium"

    if text in AMBIGUOUS_MAPPING:
        return f"Overlap/DD: {AMBIGUOUS_MAPPING[text]}"

    if text in UNSPECIFIED_SPA:
        return "SpA nicht axial/peripher spezifiziert"

    return "nicht Zielindikation"


def requires_manual_review(value: object) -> bool:
    """Markiert Verdachts-, Overlap-, DD- und unspezifische SpA-Fälle."""
    status = classify_indication_status(value)
    return status not in {"eindeutig/bestätigt", "nicht Zielindikation", "keine Indikation"}
