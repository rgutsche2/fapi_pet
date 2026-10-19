"""
===============================================================================
FAPI PET Patient Database Processing
===============================================================================

Author: Robin Gutsche
Institute: Department of Nuclear Medicine, FAU Erlangen-Nürnberg
Email: robin.gutsche@fau.de

Description:
    Reads the FAPI PET patient database and automatically creates
    disease-specific patient cohorts including follow-up information.

Version: 1.0
Date: 2026-08-05

Copyright (c) 2026 Robin Gutsche
===============================================================================
"""

#%%
import pandas as pd
from pathlib import Path

from classification import CLASSIFICATION

#%%  =============================================================================
# 1. DATEIPFADE
# =============================================================================

main_path = Path(
    "/Users/robin/SynologyDrive/_Studium_FAU/Doktorarbeit/Daten"
)

input_file = main_path / "Patientenliste_raw.xlsx"
output_file = main_path / "Indikationspruefung.xlsx"


#%% =============================================================================
# 2. PATIENTENLISTE EINLESEN
# =============================================================================

df = pd.read_excel(input_file)

x = df["Indikation"].dropna().sort_values().unique()

print(f"Patienten insgesamt: {len(df)}")
print(df.head())

#%%
def normalize_indication(value):
    """
    Vereinheitlicht eine Indikationsbezeichnung für den Vergleich.
    """

    if pd.isna(value):
        return ""

    value = str(value)
    value = value.strip()
    value = value.casefold()
    value = " ".join(value.split())

    return value

#%%
def classify_indication(value):
    """
    Ordnet eine Indikation einer Zielgruppe zu.

    Rückgabe:
    - vorgeschlagene Gruppe
    - Sicherheit
    - Kommentar
    """

    normalized_value = normalize_indication(value)

    if normalized_value == "":
        return "fehlend", "unsicher", "Keine Indikation eingetragen"

    if normalized_value in CLASSIFICATION:
        return CLASSIFICATION[normalized_value]

    return (
        "keine Zielindikation",
        "nicht relevant",
        "Keine der vier aktuell untersuchten Zielindikationen"
    )


#%%
indication_overview = (
    df["Indikation"]
    .fillna("")
    .astype(str)
    .str.strip()
    .value_counts()
    .rename_axis("Originalindikation")
    .reset_index(name="Anzahl Patienten")
)

print(f"Anzahl unterschiedlicher Indikationen: {len(indication_overview)}")
print(indication_overview.to_string(index=False))

#%%
indication_overview["Normalisierte Indikation"] = (
    indication_overview["Originalindikation"]
    .apply(normalize_indication)
)

indication_overview["Im Dictionary vorhanden"] = (
    indication_overview["Normalisierte Indikation"]
    .isin(CLASSIFICATION.keys())
)

missing_indications = indication_overview[
    indication_overview["Im Dictionary vorhanden"] == False
].copy()

print(f"Explizit klassifiziert: "
      f"{indication_overview['Im Dictionary vorhanden'].sum()}")

print(f"Noch nicht klassifiziert: {len(missing_indications)}")

print(
    missing_indications[
        ["Originalindikation", "Anzahl Patienten"]
    ].to_string(index=False)
)
#%%
output_file = main_path / "Noch_nicht_klassifizierte_Indikationen.xlsx"

missing_indications.to_excel(
    output_file,
    index=False
)

print(f"Datei erstellt: {output_file}")

#%%
reviewed_file = main_path / "nicht_relevante_Indikationen.xlsx"

reviewed_non_target_df = pd.read_excel(reviewed_file)

REVIEWED_NON_TARGET_INDICATIONS = set(
    reviewed_non_target_df["Normalisierte Indikation"]
    .dropna()
    .astype(str)
    .str.strip()
)

#%%
def classify_indication(value):
    """
    Klassifiziert eine Indikation anhand von:

    1. CLASSIFICATION:
       relevante oder möglicherweise relevante Indikationen

    2. REVIEWED_NON_TARGET_INDICATIONS:
       manuell geprüfte, nicht relevante Indikationen

    3. unbekannte neue Einträge:
       müssen erneut manuell geprüft werden
    """

    normalized_value = normalize_indication(value)

    if normalized_value == "":
        return (
            "fehlend",
            "unsicher",
            "Keine Indikation eingetragen"
        )

    # Relevante oder potenziell relevante Indikation
    if normalized_value in CLASSIFICATION:
        return CLASSIFICATION[normalized_value]

    # Von dir bereits manuell geprüft und als nicht relevant bewertet
    if normalized_value in REVIEWED_NON_TARGET_INDICATIONS:
        return (
            "keine Zielindikation",
            "sicher",
            "Manuell geprüft: keine der vier Zielindikationen"
        )

    # Neue Indikation, die in der aktuellen Liste noch nicht vorkam
    return (
        "noch nicht geprüft",
        "unsicher",
        "Neue oder bislang nicht klassifizierte Indikation"
    )

#%%
indication_overview = (
    df["Indikation"]
    .fillna("")
    .astype(str)
    .str.strip()
    .value_counts()
    .rename_axis("Originalindikation")
    .reset_index(name="Anzahl Patienten")
)

indication_overview["Normalisierte Indikation"] = (
    indication_overview["Originalindikation"]
    .apply(normalize_indication)
)

#%%
classification_results = indication_overview[
    "Originalindikation"
].apply(classify_indication)

#%%
classification_results = pd.DataFrame(
    classification_results.tolist(),
    columns=[
        "Vorgeschlagene Gruppe",
        "Sicherheit",
        "Kommentar"
    ]
)

#%%
indication_overview = pd.concat(
    [
        indication_overview,
        classification_results
    ],
    axis=1
)

#%%
uncertain_indications = indication_overview[
    indication_overview["Sicherheit"] == "unsicher"
].copy()

#%%
output_file = main_path / "Indikationspruefung_final.xlsx"

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:

    indication_overview.to_excel(
        writer,
        sheet_name="Alle_Indikationen",
        index=False
    )

    uncertain_indications.to_excel(
        writer,
        sheet_name="Fuer_Chef_pruefen",
        index=False
    )

print(f"Datei erstellt:\n{output_file}")


#%% Nach Indikation Prüfung
import pandas as pd
from pathlib import Path


main_path = Path(
    "/Users/robin/SynologyDrive/_Studium_FAU/Doktorarbeit/Daten"
)

patient_file = main_path / "Patientenliste_raw.xlsx"
classification_file = main_path / "Indikationspruefung_final_V3.xlsx"


# Ursprüngliche Patientenliste
df = pd.read_excel(patient_file)

# Fertig geprüfte Indikationsübersicht
indication_mapping = pd.read_excel(
    classification_file,
    sheet_name="Alle_Indikationen"
)

#%%
# indication_mapping["Finale_Entscheidung_Armin"] = (
#     indication_mapping["Finale_Entscheidung_Armin"]
#     .replace("", pd.NA)
# )
#
# indication_mapping["Finale Gruppe"] = (
#     indication_mapping["Finale_Entscheidung_Armin"]
#     .combine_first(
#         indication_mapping["Vorgeschlagene Gruppe"]
#     )
# )

#%%
# print(
#     indication_mapping[
#         [
#             "Originalindikation",
#             "Vorgeschlagene Gruppe",
#             "Finale_Entscheidung_Armin",
#             "Finale Gruppe"
#         ]
#     ].head(20)
# )

#%%
def normalize_indication(value):

    if pd.isna(value):
        return ""

    value = str(value)
    value = value.strip()
    value = value.casefold()
    value = " ".join(value.split())

    return value


df["Normalisierte Indikation"] = (
    df["Indikation"]
    .apply(normalize_indication)
)


indication_mapping["Normalisierte Indikation"] = (
    indication_mapping["Originalindikation"]
    .apply(normalize_indication)
)


mapping_for_patients = indication_mapping[
    [
        "Normalisierte Indikation",
        "Vorgeschlagene Gruppe"
    ]
].copy()


mapping_for_patients = (
    mapping_for_patients
    .drop_duplicates(
        subset="Normalisierte Indikation"
    )
)


print(mapping_for_patients.head())
print(f"Anzahl Mapping-Einträge: {len(mapping_for_patients)}")


df = df.merge(
    mapping_for_patients,
    on="Normalisierte Indikation",
    how="left",
    validate="many_to_one"
)


print(
    df[
        [
            "Name",
            "Indikation",
            "Vorgeschlagene Gruppe"
        ]
    ].head(20)
)


print(
    df["Vorgeschlagene Gruppe"]
    .value_counts(dropna=False)
)

#%%
patients_without_group = df[
    df["Vorgeschlagene Gruppe"].isna()
].copy()

print(
    f"Patienten ohne Gruppenzuordnung: "
    f"{len(patients_without_group)}"
)

#%%
output_file = main_path / "Patientenliste_mit_Indikationsgruppen.xlsx"

df.to_excel(
    output_file,
    index=False
)

print(f"Datei erstellt:\n{output_file}")

