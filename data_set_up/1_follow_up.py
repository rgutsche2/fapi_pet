
import pandas as pd
from pathlib import Path

#%%

main_path = Path(
    "/Users/robin/SynologyDrive/_Studium_FAU/Doktorarbeit/Daten"
)

input_file = main_path / "Patientenliste_mit_Indikationsgruppen_final.xlsx"

df = pd.read_excel(input_file)

print(f"Patienten insgesamt: {len(df)}")
print(df.columns.tolist())

# df["Studiengruppe"] = df["Gruppe"]
# df.loc[
#     df["Finale Gruppe"].isin(
#         [
#             "axSpA",
#             "pSpA",
#             "SpA",
#         ]
#     ),
#     "Studiengruppe"
# ] = "SpA"

#%%
print(
    df["Studiengruppe"]
    .value_counts(dropna=False)
)

#%%
followup_columns = [
    column
    for column in df.columns
    if str(column).startswith("Datum ")
    and "FAPI" in str(column)
]

#%%
from datetime import datetime, date
import re
import pandas as pd


def is_date(value):

    # leere Zelle
    if pd.isna(value):
        return False

    # Excel wurde von pandas direkt als Datum erkannt
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return True

    # Datum steht als Text in Excel, z.B. "19.01.26"
    if isinstance(value, str):

        value = value.strip()

        date_pattern = r"^\d{1,2}\.\d{1,2}\.\d{2,4}$"

        if re.match(date_pattern, value):
            return True

    return False


def count_followups(row):

    number_followups = 0

    for column in followup_columns:

        if is_date(row[column]):
            number_followups += 1

    return number_followups



df["Anzahl Follow-ups"] = df.apply(
    count_followups,
    axis=1
)

#%%
df["Untersuchungsstatus"] = "Nur Baseline"

df.loc[
    df["Anzahl Follow-ups"] > 0,
    "Untersuchungsstatus"
] = "Mit Follow-up"


#%%
def get_followup_dates(row):

    dates = []

    for column in followup_columns:

        value = row[column]

        if is_date(value):
            dates.append(value)

    return dates

df["Follow-up Daten"] = df.apply(
    get_followup_dates,
    axis=1
)

#%%
print(
    df[
        [
            "Name",
            "Scandatum",
            "Anzahl Follow-ups",
            "Follow-up Daten",
            "Untersuchungsstatus"
        ]
    ].to_string(index=False)
)

#%%
# Follow-up-Daten für Excel lesbar machen
df["Follow-up Daten Excel"] = df["Follow-up Daten"].apply(
    lambda dates: " | ".join(str(date) for date in dates)
)

#%%
output_file = main_path / "Patientenliste_mit_Followup_Status.xlsx"

df.to_excel(
    output_file,
    index=False
)

print(f"Datei gespeichert unter:\n{output_file}")


#%%
print(
    df["Untersuchungsstatus"]
    .value_counts()
)

print(
    df["Anzahl Follow-ups"]
    .value_counts()
    .sort_index()
)

#%%

ra = df[df["Studiengruppe"] == "RA"].copy()
psa = df[df["Studiengruppe"] == "PsA"].copy()
spa = df[df["Studiengruppe"] == "SpA"].copy()

#%%
# RA
ra_baseline = ra[ra["Untersuchungsstatus"] == "Nur Baseline"].copy()
ra_followup = ra[ra["Untersuchungsstatus"] == "Mit Follow-up"].copy()

# PsA
psa_baseline = psa[psa["Untersuchungsstatus"] == "Nur Baseline"].copy()
psa_followup = psa[psa["Untersuchungsstatus"] == "Mit Follow-up"].copy()

# SpA
spa_baseline = spa[spa["Untersuchungsstatus"] == "Nur Baseline"].copy()
spa_followup = spa[spa["Untersuchungsstatus"] == "Mit Follow-up"].copy()

#%%
print("\nRA")
print(f"Baseline:  {len(ra_baseline)}")
print(f"Follow-up: {len(ra_followup)}")

print("\nPsA")
print(f"Baseline:  {len(psa_baseline)}")
print(f"Follow-up: {len(psa_followup)}")

print("\nSpA")
print(f"Baseline:  {len(spa_baseline)}")
print(f"Follow-up: {len(spa_followup)}")

#%%
ra_file = main_path / "RA_Patienten.xlsx"

with pd.ExcelWriter(ra_file, engine="openpyxl") as writer:

    ra.to_excel(
        writer,
        sheet_name="Baseline + Follow-up",
        index=False
    )

    ra_baseline.to_excel(
        writer,
        sheet_name="Baseline",
        index=False
    )

    ra_followup.to_excel(
        writer,
        sheet_name="Follow-up",
        index=False
    )


psa_file = main_path / "PsA_Patienten.xlsx"

with pd.ExcelWriter(psa_file, engine="openpyxl") as writer:

    psa.to_excel(
        writer,
        sheet_name="Baseline + Follow-up",
        index=False
    )

    psa_baseline.to_excel(
        writer,
        sheet_name="Baseline",
        index=False
    )

    psa_followup.to_excel(
        writer,
        sheet_name="Follow-up",
        index=False
    )


spa_file = main_path / "SpA_Patienten.xlsx"

with pd.ExcelWriter(spa_file, engine="openpyxl") as writer:

    spa.to_excel(
        writer,
        sheet_name="Baseline + Follow-up",
        index=False
    )

    spa_baseline.to_excel(
        writer,
        sheet_name="Baseline",
        index=False
    )

    spa_followup.to_excel(
        writer,
        sheet_name="Follow-up",
        index=False
    )
