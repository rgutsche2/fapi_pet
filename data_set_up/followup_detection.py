"""
Erkennung und Aufbereitung von Follow-up-Untersuchungen.

Hintergrund:
In älteren Einträgen enthalten die Follow-up-Spalten häufig nur Datumswerte.
In neueren Einträgen stehen teilweise Datumswerte und Freitextnotizen in
aufeinanderfolgenden Spalten. Deshalb werden sichere Datumsangaben und
textbasierte Hinweise getrennt dokumentiert.

Wichtig:
Eine Textnotiz kann entweder zu einem vorherigen Datum gehören oder auf eine
Untersuchung mit fehlendem Datum hinweisen. Solche Fälle werden markiert,
nicht stillschweigend als sicherer zusätzlicher Scan gewertet.
"""

from __future__ import annotations

from datetime import date, datetime
from typing import Any

import pandas as pd


def is_empty(value: Any) -> bool:
    """Prüft, ob ein Zellinhalt leer ist."""
    if pd.isna(value):
        return True
    return isinstance(value, str) and value.strip() == ""


def parse_date(value: Any) -> pd.Timestamp:
    """
    Interpretiert Excel-Datumswerte und typische deutsche Textdaten.

    Nicht interpretierbare Inhalte werden als pd.NaT zurückgegeben.
    """
    if is_empty(value):
        return pd.NaT

    if isinstance(value, (pd.Timestamp, datetime, date)):
        return pd.to_datetime(value, errors="coerce")

    if isinstance(value, str):
        text = value.strip()

        # Eine reine Freitextnotiz soll nicht aggressiv als Datum interpretiert werden.
        # Deshalb werden nur Strings mit mindestens einer Ziffer geprüft.
        if not any(char.isdigit() for char in text):
            return pd.NaT

        return pd.to_datetime(text, dayfirst=True, errors="coerce")

    return pd.to_datetime(value, errors="coerce")


def classify_cell(value: Any) -> str:
    """Klassifiziert eine Zelle als leer, Datum oder Notiz."""
    if is_empty(value):
        return "leer"
    if pd.notna(parse_date(value)):
        return "Datum"
    return "Notiz"


def evaluate_followups(row: pd.Series, followup_columns: list[str]) -> list[dict[str, Any]]:
    """
    Analysiert Follow-up-Spalten von links nach rechts.

    Logik:
    - Jedes erkannte Datum erzeugt einen sicheren Follow-up-Eintrag.
    - Eine direkt danach stehende Textnotiz wird diesem Datum zugeordnet.
    - Weitere Notizen ohne zuordenbares Datum werden als unklarer Eintrag
      mit fehlendem Datum dokumentiert.
    """
    followups: list[dict[str, Any]] = []
    last_dated_followup: dict[str, Any] | None = None

    for column in followup_columns:
        value = row.get(column)

        if is_empty(value):
            continue

        parsed = parse_date(value)

        if pd.notna(parsed):
            item = {
                "datum": parsed,
                "notiz": "",
                "quelle_datum": column,
                "quelle_notiz": "",
                "status": "Datum vorhanden",
            }
            followups.append(item)
            last_dated_followup = item
            continue

        note = str(value).strip()

        # Erste Notiz nach einem Datum wird diesem Datum zugeordnet.
        if last_dated_followup is not None and not last_dated_followup["notiz"]:
            last_dated_followup["notiz"] = note
            last_dated_followup["quelle_notiz"] = column
            last_dated_followup = None
        else:
            followups.append(
                {
                    "datum": pd.NaT,
                    "notiz": note,
                    "quelle_datum": "",
                    "quelle_notiz": column,
                    "status": "Datum fehlt / manuell prüfen",
                }
            )
            last_dated_followup = None

    return followups


def add_followup_summary(
    df: pd.DataFrame,
    followup_columns: list[str],
) -> pd.DataFrame:
    """Ergänzt zusammenfassende Follow-up-Spalten."""
    result = df.copy()

    result["_Followup_Liste"] = result.apply(
        lambda row: evaluate_followups(row, followup_columns),
        axis=1,
    )

    result["Anzahl Follow-ups mit Datum"] = result["_Followup_Liste"].apply(
        lambda items: sum(pd.notna(item["datum"]) for item in items)
    )

    result["Anzahl unklare Follow-up-Hinweise"] = result["_Followup_Liste"].apply(
        lambda items: sum(pd.isna(item["datum"]) for item in items)
    )

    result["Follow-up vorhanden"] = (
        result["Anzahl Follow-ups mit Datum"]
        + result["Anzahl unklare Follow-up-Hinweise"]
        > 0
    )

    result["Manuelle Follow-up-Prüfung"] = (
        result["Anzahl unklare Follow-up-Hinweise"] > 0
    )

    return result


def build_long_followup_table(
    df: pd.DataFrame,
    id_columns: list[str],
) -> pd.DataFrame:
    """
    Erstellt eine Long-Format-Tabelle:
    eine Zeile pro Patient und Follow-up-Eintrag.
    """
    rows: list[dict[str, Any]] = []

    for _, patient in df.iterrows():
        base_data = {
            column: patient.get(column)
            for column in id_columns
            if column in df.columns
        }

        followups = patient.get("_Followup_Liste", [])

        if not followups:
            rows.append(
                {
                    **base_data,
                    "Follow-up Nummer": 0,
                    "Follow-up Datum": pd.NaT,
                    "Follow-up Notiz": "",
                    "Follow-up Status": "Kein Follow-up dokumentiert",
                    "Quellspalte Datum": "",
                    "Quellspalte Notiz": "",
                }
            )
            continue

        for number, item in enumerate(followups, start=1):
            rows.append(
                {
                    **base_data,
                    "Follow-up Nummer": number,
                    "Follow-up Datum": item["datum"],
                    "Follow-up Notiz": item["notiz"],
                    "Follow-up Status": item["status"],
                    "Quellspalte Datum": item["quelle_datum"],
                    "Quellspalte Notiz": item["quelle_notiz"],
                }
            )

    return pd.DataFrame(rows)
