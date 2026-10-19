"""
Zentrale Verarbeitungspipeline für die FAPI-PET-Patientenliste.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

from data_set_up.config import (
    BASELINE_DATE_COLUMN,
    BIRTHDATE_COLUMN,
    EXPORT_INDICATION_OVERVIEW,
    EXPORT_REVIEW_TABLES,
    FOLLOWUP_COLUMN_PREFIX,
    ID_COLUMN,
    INDICATION_COLUMN,
    INPUT_FILE,
    INPUT_SHEET,
    NAME_COLUMN,
    OUTPUT_FILE,
    TARGET_GROUPS,
)
from data_set_up.excel_output import write_excel_file
from data_set_up.followup_detection import add_followup_summary, build_long_followup_table
from data_set_up.indication_mapping import (
    classify_indication,
    classify_indication_status,
    requires_manual_review,
)


def validate_input_file(input_file: Path) -> None:
    """Prüft, ob die Eingabedatei vorhanden ist."""
    if not input_file.exists():
        raise FileNotFoundError(
            "\nDie Eingabedatei wurde nicht gefunden:\n"
            f"{input_file}\n\n"
            "Lege 'Patientenliste_raw.xlsx' in denselben Ordner wie main.py "
            "oder passe INPUT_FILE in config.py an."
        )


def load_patient_data(input_file: Path) -> pd.DataFrame:
    """Liest die Excel-Patientenliste ein."""
    validate_input_file(input_file)

    try:
        df = pd.read_excel(input_file, sheet_name=INPUT_SHEET)
    except ImportError as exc:
        raise RuntimeError(
            "Zum Lesen von Excel-Dateien wird openpyxl benötigt. "
            "Installiere die Abhängigkeiten mit: pip install -r requirements.txt"
        ) from exc

    df.columns = [str(column).strip() for column in df.columns]
    return df


def validate_required_columns(df: pd.DataFrame) -> None:
    """Prüft die minimal benötigten Spalten."""
    required = [NAME_COLUMN, BASELINE_DATE_COLUMN, INDICATION_COLUMN]
    missing = [column for column in required if column not in df.columns]

    if missing:
        raise KeyError(
            "Folgende benötigte Spalten fehlen in der Eingabedatei: "
            + ", ".join(missing)
            + "\n\nVorhandene Spalten:\n"
            + "\n".join(df.columns.astype(str))
        )


def find_followup_columns(df: pd.DataFrame) -> list[str]:
    """
    Erkennt Follow-up-Spalten automatisch.

    Erwartet Spalten wie:
    'Datum 2. FAPI PET/CT', 'Datum 3. FAPI PET/CT', ...
    """
    columns = [
        column
        for column in df.columns
        if str(column).strip().lower().startswith(FOLLOWUP_COLUMN_PREFIX.lower())
        and column != BASELINE_DATE_COLUMN
    ]

    if not columns:
        print("WARNUNG: Es wurden keine Follow-up-Spalten erkannt.")

    return columns


def prepare_data(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """Bereinigt Datumsfelder, klassifiziert Indikationen und Follow-ups."""
    result = df.copy()
# R: Anpassen von Datum Format
    for column in [BIRTHDATE_COLUMN, BASELINE_DATE_COLUMN]:
        if column in result.columns:
            result[column] = pd.to_datetime(
                result[column],
                dayfirst=True,
                errors="coerce",
            )

    result["Indikation Original"] = result[INDICATION_COLUMN]
    result["Gruppe"] = result[INDICATION_COLUMN].apply(classify_indication)
    result["Zuordnungsstatus"] = result[INDICATION_COLUMN].apply(
        classify_indication_status
    )
    result["Indikation manuell prüfen"] = result[INDICATION_COLUMN].apply(
        requires_manual_review
    )
    result["Indikation zugeordnet"] = result["Gruppe"].notna()

    followup_columns = find_followup_columns(result)
    result = add_followup_summary(result, followup_columns)

    return result, followup_columns


def create_indication_overview(df: pd.DataFrame) -> pd.DataFrame:
    """Erstellt eine Übersicht aller vorkommenden Indikationen."""
    overview = (
        df.groupby(["Indikation Original", "Gruppe"], dropna=False)
        .size()
        .reset_index(name="Anzahl Patienten")
        .sort_values(
            ["Gruppe", "Anzahl Patienten", "Indikation Original"],
            ascending=[True, False, True],
            na_position="last",
        )
    )

    overview["Prüfstatus"] = overview["Gruppe"].apply(
        lambda value: "zugeordnet" if pd.notna(value) else "noch nicht zugeordnet"
    )

    return overview


def build_output_sheets(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Erzeugt alle Tabellenblätter für die Excel-Ausgabe."""
    sheets: dict[str, pd.DataFrame] = {}

    preferred_columns = [
        ID_COLUMN,
        NAME_COLUMN,
        BIRTHDATE_COLUMN,
        BASELINE_DATE_COLUMN,
        INDICATION_COLUMN,
        "Indikation Original",
        "Gruppe",
        "Zuordnungsstatus",
        "Indikation manuell prüfen",
        "Follow-up vorhanden",
        "Anzahl Follow-ups mit Datum",
        "Anzahl unklare Follow-up-Hinweise",
        "Manuelle Follow-up-Prüfung",
    ]

    existing_preferred = [
        column for column in preferred_columns if column in df.columns
    ]
    remaining = [
        column
        for column in df.columns
        if column not in existing_preferred and not column.startswith("_")
    ]
    ordered_columns = existing_preferred + remaining

    id_columns = [
        column
        for column in [
            ID_COLUMN,
            NAME_COLUMN,
            BIRTHDATE_COLUMN,
            BASELINE_DATE_COLUMN,
            "Indikation Original",
            "Gruppe",
        ]
        if column in df.columns
    ]

    for group in TARGET_GROUPS:
        group_df = (
            df.loc[df["Gruppe"] == group, ordered_columns]
            .sort_values(
                [BASELINE_DATE_COLUMN, NAME_COLUMN],
                ascending=[True, True],
                na_position="last",
            )
            .reset_index(drop=True)
        )

        sheets[f"{group}_Alle"] = group_df

        sheets[f"{group}_Followups"] = group_df.loc[
            group_df["Follow-up vorhanden"]
        ].reset_index(drop=True)

        group_long = build_long_followup_table(
            df.loc[df["Gruppe"] == group],
            id_columns=id_columns,
        )
        sheets[f"{group}_Followup_Long"] = group_long

        if EXPORT_REVIEW_TABLES:
            sheets[f"{group}_Pruefung"] = group_df.loc[
                group_df["Manuelle Follow-up-Prüfung"]
            ].reset_index(drop=True)

    unmapped = (
        df.loc[df["Gruppe"].isna(), ordered_columns]
        .sort_values(
            ["Indikation Original", NAME_COLUMN],
            na_position="last",
        )
        .reset_index(drop=True)
    )
    sheets["Nicht_zugeordnet"] = unmapped

    if EXPORT_INDICATION_OVERVIEW:
        sheets["Indikationen_Uebersicht"] = create_indication_overview(df)

    return sheets


def print_summary(df: pd.DataFrame, output_file: Path) -> None:
    """Gibt eine kompakte Zusammenfassung in der Konsole aus."""
    print("\n" + "=" * 60)
    print("FAPI PET PATIENT DATABASE PROCESSING")
    print("=" * 60)
    print(f"Patienten insgesamt: {len(df)}")

    for group in TARGET_GROUPS:
        group_df = df.loc[df["Gruppe"] == group]
        followup_count = int(group_df["Follow-up vorhanden"].sum())
        review_count = int(group_df["Manuelle Follow-up-Prüfung"].sum())

        print(
            f"{group}: {len(group_df)} Patienten | "
            f"{followup_count} mit Follow-up-Hinweis | "
            f"{review_count} manuell prüfen"
        )

    print(f"Nicht zugeordnet: {int(df['Gruppe'].isna().sum())}")
    print("-" * 60)
    print(f"Ausgabe erstellt:\n{output_file}")
    print("=" * 60 + "\n")


def run_pipeline() -> None:
    """Führt die komplette Verarbeitung aus."""
    try:
        print(f"Lese Eingabedatei:\n{INPUT_FILE}\n")

        df = load_patient_data(INPUT_FILE)
        validate_required_columns(df)

        prepared_df, followup_columns = prepare_data(df)

        print("Erkannte Follow-up-Spalten:")
        for column in followup_columns:
            print(f"  - {column}")

        sheets = build_output_sheets(prepared_df)
        write_excel_file(OUTPUT_FILE, sheets)
        print_summary(prepared_df, OUTPUT_FILE)

    except Exception as exc:
        print("\nFEHLER BEI DER VERARBEITUNG")
        print("-" * 60)
        print(str(exc))
        print("-" * 60)
        sys.exit(1)
