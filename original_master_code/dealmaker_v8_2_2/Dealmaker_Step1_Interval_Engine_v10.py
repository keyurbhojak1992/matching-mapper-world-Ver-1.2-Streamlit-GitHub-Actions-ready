#!/usr/bin/env python3
"""Dealmaker Step 1 - interval engine.

Creates/updates one control workbook that serves two purposes:
1. Human summary of real deal buyers and their deal intervals.
2. Machine-ready Stock_Fetcher_Input sheet containing only interval-eligible parties.

Business-safe defaults:
- Only Type == SALE is treated as completed business.
- Stones are counted by unique Serial No per party and business date.
- A real deal contains at least MIN_DEAL_STONES unique stones in one day or a
  continuous two-day bucket.
- Median interval is the primary cadence measure.
- Automatic interval eligibility requires repeat deal history.
- BID and action activity are kept as intent signals, never as completed sales.
"""

from __future__ import annotations

import argparse
import hashlib
import math
import os
import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable

import numpy as np
import pandas as pd
import dealmaker_quality_v6 as quality_domain
from dealmaker_quality_v6 import quality_score as shared_quality_score
from dealmaker_io_v2 import load_manual_mapping, read_dataset_table, serial_key_series
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

QUALITY_ORDERS = quality_domain.QUALITY_ORDERS
SCORE_MAPS = quality_domain.QUALITY_MAPS
REV_SCORE_MAPS = quality_domain.QUALITY_REVERSE

def configure_quality(control_path):
    global QUALITY_ORDERS, SCORE_MAPS, REV_SCORE_MAPS
    cfg = quality_domain.configure_quality(control_path)
    QUALITY_ORDERS = cfg.orders
    SCORE_MAPS = cfg.score_maps
    REV_SCORE_MAPS = cfg.reverse_maps
    globals()["SHADE_ORDER"] = QUALITY_ORDERS["Shade"]
    globals()["MILKY_ORDER"] = QUALITY_ORDERS["Milky"]
    globals()["OPEN_ORDER"] = QUALITY_ORDERS["Open"]
    globals()["BLK_ORDER"] = QUALITY_ORDERS["Blk"]
    return cfg


VERSION = "8.2.2"

DEFAULT_SETTINGS = [
    ("MIN_DEAL_STONES", 6, "Minimum unique SALE stones in a real deal bucket."),
    ("MAX_DEAL_DAYS", 2, "Maximum continuous calendar days in one deal bucket. This version supports 1 or 2."),
    ("MIN_DEAL_BUCKETS_FOR_MAKER", 3, "Minimum completed deal buckets to call a party a DEAL MAKER."),
    ("MIN_INTERVALS_FOR_AUTO", 2, "Minimum measured intervals needed for automatic due-date output."),
    ("DUE_SOON_RATIO", 0.80, "DUE SOON begins at this fraction of the median interval."),
    ("DUE_NOW_RATIO", 1.00, "DUE NOW begins at this fraction of the median interval."),
    ("OVERDUE_RATIO", 1.30, "OVERDUE begins above this fraction of the median interval."),
    ("INCLUDE_DUE_SOON", "NO", "YES includes DUE SOON parties in Stock_Fetcher_Input."),
    ("REQUIRE_RECENT_ACTION", "NO", "YES requires an action within RECENT_ACTION_DAYS for automatic inclusion."),
    ("RECENT_ACTION_DAYS", 90, "Action recency window used only when REQUIRE_RECENT_ACTION is YES."),
    ("PROFILE_LOOKBACK_DAYS", 730, "SALE history lookback for preference profiles. Blank or 0 means all history."),
    ("MIN_PROFILE_SALE_STONES", 1, "Minimum SALE stones for a Party + Supertag profile row."),
    ("AS_OF_DATE", "", "Optional DD-MM-YYYY override. Blank uses latest valid Sales/Action date."),
    ("SALE_TYPE_VALUES", "SALE", "Comma-separated completed-sale Type values."),
    ("BID_TYPE_VALUES", "BID", "Comma-separated bid/intent Type values."),
    ("UNKNOWN_QUALITY_POLICY", "FLAG", "FLAG sends unknown-quality stock rows to review instead of the final deal file."),
    ("MIN_STOCK_MATCHES_PER_PARTY", 2, "Minimum final matched stock rows required to create a party sheet."),
    ("INCLUDE_SECONDARY_MATCHES", "YES", "YES includes stock within maximum limits but above majority preference."),
    ("PAIR_CASCADE_EXCLUSION", "YES", "YES removes both independently matched pair members when one is excluded."),
]

ACTION_WEIGHTS = {
    "OFFER": 6.0,
    "CART": 4.0,
    "WISHLIST": 3.0,
    "DETAILS": 1.0,
    "VIEW": 0.5,
}

SHADE_ORDER = QUALITY_ORDERS["Shade"]
MILKY_ORDER = QUALITY_ORDERS["Milky"]
OPEN_ORDER = QUALITY_ORDERS["Open"]
BLK_ORDER = QUALITY_ORDERS["Blk"]

STD_CARAT_ORDER = [
    (0.30, 0.39, "0.30-0.39"),
    (0.40, 0.49, "0.40-0.49"),
    (0.50, 0.59, "0.50-0.59"),
    (0.60, 0.69, "0.60-0.69"),
    (0.70, 0.79, "0.70-0.79"),
    (0.80, 0.89, "0.80-0.89"),
    (0.90, 0.99, "0.90-0.99"),
    (1.00, 1.09, "1.00-1.09"),
    (1.10, 1.19, "1.10-1.19"),
    (1.20, 1.24, "1.20-1.24"),
    (1.25, 1.49, "1.25-1.49"),
    (1.50, 1.69, "1.50-1.69"),
    (1.70, 1.99, "1.70-1.99"),
    (2.00, 2.49, "2.00-2.49"),
    (2.50, 2.99, "2.50-2.99"),
    (3.00, 3.49, "3.00-3.49"),
    (3.50, 3.99, "3.50-3.99"),
    (4.00, 4.49, "4.00-4.49"),
    (4.50, 4.99, "4.50-4.99"),
    (5.00, 7.99, "5.00-7.99"),
    (8.00, 14.99, "8.00-14.99"),
    (15.00, 999.0, "15+"),
]

COLUMN_ALIASES = {
    "Serial No": ["Serial No", "Stone ID", "SerialNo"],
    "Sold Party": ["Sold Party", "Party", "Customer", "Party Name"],
    "Date": ["Date", "Business Date", "Sale Date", "Action Date"],
    "Type": ["Type", "Transaction Type", "Action Type"],
    "Sales_Person": ["Sales_Person", "SALE_PERSON", "Sales Person", "Salesperson"],
    "Country": ["Country"],
    "City": ["City"],
    "Shp": ["Shp", "Shape"],
    "Carat": ["Carat"],
    "Carat_grp": ["Carat_grp", "carat_grp", "Carat Group"],
    "Col": ["Col", "Color", "Color name"],
    "Qua": ["Qua", "Clarity"],
    "Flo": ["Flo", "Flou.", "Fluorescence"],
    "Lab": ["Lab"],
    "Shade": ["Shade"],
    "Milky": ["Milky"],
    "TOIN_NAME": ["TOIN_NAME", "Table Open"],
    "SOIN_NAME": ["SOIN_NAME", "Side Open"],
    "GIRDLE_OPEN_NAME": ["GIRDLE_OPEN_NAME", "Girdle Open"],
    "PAV_OPEN_NAME": ["PAV_OPEN_NAME", "Pavilion Open"],
    "TBINC_NAME": ["TBINC_NAME", "Table Black", "Table Black Incl."],
    "SBINC_NAME": ["SBINC_NAME", "Side Black", "Side Black Incl."],
    "Disc%": ["Disc%", "Dis%", "Discount", "Reserve%"],
}

GENERATED_SHEETS = [
    "Dashboard", "Run_Summary", "Party_Deal_Summary", "Deal_Buckets",
    "Deal_Interval_History", "Eligible_Parties", "Stock_Fetcher_Input",
    "Matcher_Profile_All", "Data_Quality", "Code_Audit", "Legacy_Input_Audit",
]
INPUT_SHEETS = ["Settings", "Party_Overrides", "Removed_Parties", "Exclusion_Rules", "User_Guide"]


@dataclass
class Inputs:
    sales: pd.DataFrame
    actions: pd.DataFrame
    settings: dict[str, Any]
    overrides: pd.DataFrame
    removed: pd.DataFrame
    exclusions: pd.DataFrame
    as_of: pd.Timestamp


def clean_text(value: Any, default: str = "") -> str:
    if pd.isna(value):
        return default
    return str(value).strip()


def normalize_upper(series: pd.Series, default: str = "NONE") -> pd.Series:
    return series.fillna(default).astype(str).str.strip().str.upper().replace("", default)


def standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    lookup: dict[str, str] = {}
    for standard, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            lookup[str(alias).strip().lower()] = standard
    df.columns = [lookup.get(str(c).strip().lower(), str(c).strip()) for c in df.columns]
    return df


def normalize_date(series: pd.Series) -> pd.Series:
    raw = series.copy()
    p1 = pd.to_datetime(raw.astype(str).str.strip(), format="%d-%m-%Y", errors="coerce")
    p2 = pd.to_datetime(raw.astype(str).str.strip(), format="%d/%m/%Y", errors="coerce")
    p3 = pd.to_datetime(raw, errors="coerce", dayfirst=True)
    return p1.fillna(p2).fillna(p3).dt.normalize()


def get_carat_group(value: Any) -> str:
    try:
        x = float(value)
    except Exception:
        return "OTHER"
    for low, high, label in STD_CARAT_ORDER:
        if low <= x <= high:
            return label
    return "OTHER"


def score_value(value: Any, dimension: str) -> float:
    return shared_quality_score(value, dimension, blank_as_none=False)


def row_max_score(row: pd.Series, columns: list[str], dimension: str) -> float:
    values = [score_value(row.get(c), dimension) for c in columns]
    valid = [v for v in values if not pd.isna(v)]
    return max(valid) if valid else 0.0


def majority_score(series: pd.Series) -> float:
    s = pd.to_numeric(series, errors="coerce").dropna()
    if s.empty:
        return np.nan
    modes = s.mode()
    return float(modes.min()) if not modes.empty else float(s.median())


def score_to_text(value: Any, dimension: str) -> str:
    if pd.isna(value):
        return "UNKNOWN"
    return REV_SCORE_MAPS[dimension].get(int(value), "UNKNOWN")


def first_nonblank(series: pd.Series) -> str:
    for value in series:
        text = clean_text(value)
        if text:
            return text
    return ""


def parse_csv_set(value: Any) -> set[str]:
    return {x.strip().upper() for x in str(value).split(",") if x.strip()}


def yes(value: Any) -> bool:
    return str(value).strip().upper() in {"YES", "Y", "TRUE", "1"}


def load_control_inputs(control_path: Path) -> tuple[dict[str, Any], pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    settings = {key: value for key, value, _ in DEFAULT_SETTINGS}
    overrides = pd.DataFrame(columns=["Sold Party", "Override", "Manual_Interval_Days", "Notes"])
    removed = pd.DataFrame(columns=["Sold Party", "Reason"])
    exclusions = pd.DataFrame(columns=[
        "Sold Party", "Shp", "Carat", "Color Name", "Qua", "Cut", "Pol", "Sym", "Flo", "Lab",
        "Disc%", "Shade", "Milky", "TOIN_NAME", "SOIN_NAME", "GIRDLE_OPEN_NAME",
        "PAV_OPEN_NAME", "TBINC_NAME", "SBINC_NAME", "Remarks",
    ])
    if not control_path.exists():
        return settings, overrides, removed, exclusions
    try:
        xls = pd.ExcelFile(control_path)
        def read_user_sheet(sheet_name: str, required: set[str]) -> pd.DataFrame:
            first = pd.read_excel(control_path, sheet_name=sheet_name)
            if required.issubset(set(first.columns)):
                return first
            titled = pd.read_excel(control_path, sheet_name=sheet_name, header=2)
            return titled

        if "Settings" in xls.sheet_names:
            s = read_user_sheet("Settings", {"Setting", "Value"})
            if {"Setting", "Value"}.issubset(s.columns):
                for _, row in s.iterrows():
                    key = clean_text(row["Setting"])
                    if key:
                        settings[key] = row["Value"]
        if "Party_Overrides" in xls.sheet_names:
            overrides = read_user_sheet("Party_Overrides", {"Sold Party", "Override"})
        if "Removed_Parties" in xls.sheet_names:
            removed = read_user_sheet("Removed_Parties", {"Sold Party"})
        if "Exclusion_Rules" in xls.sheet_names:
            exclusions = read_user_sheet("Exclusion_Rules", {"Sold Party"})
    except Exception as exc:
        raise RuntimeError(f"Could not read control workbook {control_path}: {exc}") from exc
    return settings, overrides, removed, exclusions


def prepare_dataset(
    path: Path,
    dataset_name: str,
    quality_log: list[dict[str, Any]],
    sheet_name: str | None = None,
    header_row: int | None = None,
    manual_mapping: dict[str, str] | None = None,
) -> pd.DataFrame:
    required = ["Serial No", "Sold Party", "Date", "Type"]
    raw, input_meta = read_dataset_table(
        path, COLUMN_ALIASES, required, dataset_name,
        sheet_name=sheet_name, header_row=header_row, manual_mapping=manual_mapping,
    )
    df = standardize_columns(raw)
    for col in ["Sales_Person", "Country", "City", "Shp", "Col", "Qua", "Flo", "Lab", "Shade", "Milky"]:
        if col not in df.columns:
            df[col] = ""
    for col in ["TOIN_NAME", "SOIN_NAME", "GIRDLE_OPEN_NAME", "PAV_OPEN_NAME", "TBINC_NAME", "SBINC_NAME"]:
        if col not in df.columns:
            df[col] = ""
    if "Carat" not in df.columns:
        df["Carat"] = np.nan
    if "Carat_grp" not in df.columns:
        df["Carat_grp"] = df["Carat"].apply(get_carat_group)
    if "Disc%" not in df.columns:
        df["Disc%"] = np.nan

    df["Raw_Date"] = df["Date"]
    df["Date"] = normalize_date(df["Date"])
    df["Sold Party"] = df["Sold Party"].fillna("").astype(str).str.strip()
    df["Type"] = normalize_upper(df["Type"], "UNKNOWN")
    df["Serial_Key"] = serial_key_series(df["Serial No"])
    df["Sales_Person"] = df["Sales_Person"].fillna("").astype(str).str.strip()
    df["Country"] = df["Country"].fillna("").astype(str).str.strip()
    df["City"] = df["City"].fillna("").astype(str).str.strip()

    for col in ["Shp", "Col", "Qua", "Flo", "Lab"]:
        df[col] = normalize_upper(df[col], "UNKNOWN")
    quality_dimensions = {
        "Shade": "Shade", "Milky": "Milky",
        "TOIN_NAME": "Open", "SOIN_NAME": "Open", "GIRDLE_OPEN_NAME": "Open", "PAV_OPEN_NAME": "Open",
        "TBINC_NAME": "Blk", "SBINC_NAME": "Blk",
    }
    for col, dim in quality_dimensions.items():
        df[f"Raw_{col}"] = df[col]
        df[col] = df[col].map(lambda x: quality_domain.canonical_quality(x, dim, blank_as_none=False))
    df["Flo"] = df["Flo"].replace({"VSL": "FNT", "SL": "FNT"})
    df["Carat"] = pd.to_numeric(df["Carat"], errors="coerce")
    # Rebuild Carat_grp from numeric Carat so legacy/pre-existing group labels
    # cannot survive after the standard carat ranges are changed. If Carat is
    # unavailable, preserve the supplied Carat_grp value as a fallback.
    supplied_carat_grp = df["Carat_grp"].fillna("").astype(str).str.strip()
    derived_carat_grp = df["Carat"].apply(get_carat_group)
    df["Carat_grp"] = derived_carat_grp.where(df["Carat"].notna(), supplied_carat_grp)
    df["Disc%"] = pd.to_numeric(df["Disc%"], errors="coerce")

    df["Supertag"] = (
        df["Shp"] + "+" + df["Carat_grp"] + "+" + df["Col"] + "+" +
        df["Qua"] + "+" + df["Flo"] + "+" + df["Lab"]
    )
    df["Shade_Score"] = df["Shade"].apply(lambda x: score_value(x, "Shade"))
    df["Milky_Score"] = df["Milky"].apply(lambda x: score_value(x, "Milky"))
    df["Open_Score"] = df.apply(
        lambda r: row_max_score(r, ["TOIN_NAME", "SOIN_NAME", "GIRDLE_OPEN_NAME", "PAV_OPEN_NAME"], "Open"), axis=1
    )
    df["Blk_Score"] = df.apply(
        lambda r: row_max_score(r, ["TBINC_NAME", "SBINC_NAME"], "Blk"), axis=1
    )

    quality_log.extend([
        {"Dataset": dataset_name, "Check": "Selected worksheet", "Value": input_meta.sheet_name or "Delimited file", "Severity": "INFO"},
        {"Dataset": dataset_name, "Check": "Detected header row", "Value": input_meta.header_row_zero_based + 1, "Severity": "INFO"},
        {"Dataset": dataset_name, "Check": "Input columns", "Value": input_meta.column_count, "Severity": "INFO"},
        {"Dataset": dataset_name, "Check": "Input rows", "Value": len(df), "Severity": "INFO"},
        {"Dataset": dataset_name, "Check": "Invalid/blank dates", "Value": int(df["Date"].isna().sum()), "Severity": "ERROR" if df["Date"].isna().any() else "OK"},
        {"Dataset": dataset_name, "Check": "Blank parties", "Value": int(df["Sold Party"].eq("").sum()), "Severity": "ERROR" if df["Sold Party"].eq("").any() else "OK"},
        {"Dataset": dataset_name, "Check": "Blank serials", "Value": int(df["Serial_Key"].eq("").sum()), "Severity": "WARNING" if df["Serial_Key"].eq("").any() else "OK"},
        {"Dataset": dataset_name, "Check": "Unknown shade values", "Value": int(df["Shade_Score"].isna().sum()), "Severity": "WARNING" if df["Shade_Score"].isna().any() else "OK"},
        {"Dataset": dataset_name, "Check": "Unknown milky values", "Value": int(df["Milky_Score"].isna().sum()), "Severity": "WARNING" if df["Milky_Score"].isna().any() else "OK"},
    ])
    return df[df["Date"].notna() & df["Sold Party"].ne("")].copy()


def determine_as_of(settings: dict[str, Any], sales: pd.DataFrame, actions: pd.DataFrame) -> pd.Timestamp:
    manual = clean_text(settings.get("AS_OF_DATE"))
    if manual:
        parsed = normalize_date(pd.Series([manual])).iloc[0]
        if pd.isna(parsed):
            raise ValueError(f"Invalid AS_OF_DATE: {manual}. Use DD-MM-YYYY.")
        return parsed
    dates = []
    if not sales.empty:
        dates.append(sales["Date"].max())
    if not actions.empty:
        dates.append(actions["Date"].max())
    if not dates:
        return pd.Timestamp(date.today())
    return max(dates)


def dedupe_business_rows(df: pd.DataFrame, dataset_name: str, quality_log: list[dict[str, Any]]) -> pd.DataFrame:
    before = len(df)
    with_serial = df["Serial_Key"].ne("")
    deduped_serial = df[with_serial].drop_duplicates(["Sold Party", "Date", "Type", "Serial_Key"], keep="last")
    no_serial = df[~with_serial].copy()
    out = pd.concat([deduped_serial, no_serial], ignore_index=True)
    quality_log.append({
        "Dataset": dataset_name,
        "Check": "Duplicate Party + Date + Type + Serial rows removed",
        "Value": before - len(out),
        "Severity": "WARNING" if before != len(out) else "OK",
    })
    return out


def build_deal_buckets(sales_only: pd.DataFrame, min_stones: int, max_days: int) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Build strongest non-overlapping 1/2-day completed-sale buckets.

    The production implementation uses a linear back-pointer dynamic program
    and a lightweight day-index builder. This preserves the confirmed v7
    objective while avoiding expensive object-heavy groupby aggregation on
    full production histories.
    """
    if max_days not in (1, 2):
        raise ValueError("MAX_DEAL_DAYS must be 1 or 2 in this version.")
    if sales_only.empty:
        return pd.DataFrame(), pd.DataFrame()

    # Build one record per party/date in a single sorted pass. All production
    # serials are expected to be present, but the row-count fallback preserves
    # the original behavior for blank-serial legacy rows.
    work = sales_only[["Sold Party", "Date", "Serial_Key", "Supertag", "Sales_Person", "Country", "City"]].sort_values(
        ["Sold Party", "Date"], kind="mergesort"
    )
    party_days: dict[str, list[dict[str, Any]]] = {}
    current_key: tuple[Any, Any] | None = None
    current_serials: set[str] = set()
    current_supertags: set[str] = set()
    current_rows = 0
    current_sales_person = ""
    current_country = ""
    current_city = ""

    def first_text(current: str, value: Any) -> str:
        if current:
            return current
        text = "" if value is None or pd.isna(value) else str(value).strip()
        return text

    def flush_day() -> None:
        nonlocal current_key, current_serials, current_supertags, current_rows, current_sales_person, current_country, current_city
        if current_key is None:
            return
        party, business_date = current_key
        party_days.setdefault(party, []).append({
            "Date": business_date,
            "Day_Rows": current_rows,
            "Day_Stones": len(current_serials) if current_serials else current_rows,
            "Day_Serial_Set": frozenset(current_serials),
            "Day_Supertag_Set": frozenset(current_supertags),
            "Sales_Person": current_sales_person,
            "Country": current_country,
            "City": current_city,
        })

    for party, business_date, serial, supertag, sales_person, country, city in work.itertuples(index=False, name=None):
        key = (party, business_date)
        if current_key is not None and key != current_key:
            flush_day()
            current_serials = set()
            current_supertags = set()
            current_rows = 0
            current_sales_person = ""
            current_country = ""
            current_city = ""
        current_key = key
        if serial:
            current_serials.add(str(serial))
        if supertag:
            current_supertags.add(str(supertag))
        current_sales_person = first_text(current_sales_person, sales_person)
        current_country = first_text(current_country, country)
        current_city = first_text(current_city, city)
        current_rows += 1
    flush_day()

    deals: list[dict[str, Any]] = []
    membership: list[dict[str, Any]] = []

    for party, rows in party_days.items():
        n = len(rows)
        best_stones = [0] * (n + 2)
        best_buckets = [0] * (n + 2)
        choice = [0] * n  # 0=skip, 1=one-day, 2=two-day

        for i in range(n - 1, -1, -1):
            chosen_stones = best_stones[i + 1]
            chosen_buckets = best_buckets[i + 1]
            chosen = 0

            stones1 = int(rows[i]["Day_Stones"])
            if stones1 >= min_stones:
                candidate_stones = stones1 + best_stones[i + 1]
                candidate_buckets = 1 + best_buckets[i + 1]
                if (candidate_stones, -candidate_buckets) > (chosen_stones, -chosen_buckets):
                    chosen_stones = candidate_stones
                    chosen_buckets = candidate_buckets
                    chosen = 1

            if max_days >= 2 and i + 1 < n and (rows[i + 1]["Date"] - rows[i]["Date"]).days == 1:
                set1 = rows[i]["Day_Serial_Set"]
                set2 = rows[i + 1]["Day_Serial_Set"]
                stones2 = len(set1.union(set2)) if (set1 or set2) else stones1 + int(rows[i + 1]["Day_Stones"])
                if stones2 >= min_stones:
                    candidate_stones = stones2 + best_stones[i + 2]
                    candidate_buckets = 1 + best_buckets[i + 2]
                    if (candidate_stones, -candidate_buckets) > (chosen_stones, -chosen_buckets):
                        chosen_stones = candidate_stones
                        chosen_buckets = candidate_buckets
                        chosen = 2

            best_stones[i] = chosen_stones
            best_buckets[i] = chosen_buckets
            choice[i] = chosen

        selected_ranges: list[tuple[int, int]] = []
        i = 0
        while i < n:
            if choice[i] == 1:
                selected_ranges.append((i, i))
                i += 1
            elif choice[i] == 2:
                selected_ranges.append((i, i + 1))
                i += 2
            else:
                i += 1

        for start_i, end_i in selected_ranges:
            start = rows[start_i]
            end = rows[end_i]
            selected_days = rows[start_i:end_i + 1]
            serials: set[str] = set()
            supertags: set[str] = set()
            row_total = 0
            sales_person = ""
            country = ""
            city = ""
            for day in selected_days:
                serials.update(day["Day_Serial_Set"])
                supertags.update(day["Day_Supertag_Set"])
                row_total += int(day["Day_Rows"])
                sales_person = first_text(sales_person, day["Sales_Person"])
                country = first_text(country, day["Country"])
                city = first_text(city, day["City"])
            serial_count = len(serials) if serials else row_total
            deals.append({
                "Sold Party": party,
                "Deal_Start_Date": start["Date"],
                "Deal_End_Date": end["Date"],
                "Deal_Days": int((end["Date"] - start["Date"]).days + 1),
                "Deal_Total_Stones": int(serial_count),
                "Deal_Unique_Supertags": int(len(supertags)),
                "Sales_Person": sales_person,
                "Country": country,
                "City": city,
            })
            for business_date in pd.date_range(start["Date"], end["Date"], freq="D"):
                membership.append({"Sold Party": party, "Date": business_date.normalize()})

    deals_df = pd.DataFrame(deals)
    if deals_df.empty:
        return deals_df, pd.DataFrame(membership)
    deals_df = deals_df.sort_values(["Sold Party", "Deal_Start_Date", "Deal_End_Date"]).reset_index(drop=True)
    deals_df["Party_Deal_Number"] = deals_df.groupby("Sold Party").cumcount() + 1
    # Globally unique and deterministic.  The v4 eight-character party prefix
    # could collide across different customers and contaminate recent-pattern
    # membership.  The party hash, date, and party sequence form a stable key.
    def make_deal_id(row: pd.Series) -> str:
        party = str(row["Sold Party"]).strip().upper()
        party_hash = hashlib.sha1(party.encode("utf-8")).hexdigest()[:12].upper()
        deal_date = pd.Timestamp(row["Deal_Start_Date"]).strftime("%Y%m%d")
        return f"DL-{party_hash}-{deal_date}-{int(row['Party_Deal_Number']):03d}"

    deals_df["Deal_ID"] = deals_df.apply(make_deal_id, axis=1)
    deals_df["Previous_Deal_Start_Date"] = deals_df.groupby("Sold Party")["Deal_Start_Date"].shift(1)
    deals_df["Interval_Days"] = (
        deals_df["Deal_Start_Date"] - deals_df["Previous_Deal_Start_Date"]
    ).dt.days
    deals_df["Deal_Definition"] = f">= {min_stones} unique SALE stones in 1-{max_days} continuous day(s)"
    return deals_df, pd.DataFrame(membership).drop_duplicates()

def aggregate_actions(actions: pd.DataFrame, as_of: pd.Timestamp, recent_days: int) -> pd.DataFrame:
    if actions.empty:
        return pd.DataFrame(columns=["Sold Party"])
    a = actions.copy()
    a["Action_Weight"] = a["Type"].map(ACTION_WEIGHTS).fillna(0.25)
    a["Is_Recent_Action"] = a["Date"] >= (as_of - pd.Timedelta(days=recent_days))
    summary = a.groupby("Sold Party").agg(
        Last_Action_Date=("Date", "max"),
        Total_Action_Count=("Type", "size"),
        Recent_Action_Count=("Is_Recent_Action", "sum"),
        Action_Weighted_Score=("Action_Weight", "sum"),
    ).reset_index()
    counts = a.groupby(["Sold Party", "Type"]).size().unstack(fill_value=0).reset_index()
    counts = counts.rename(columns={c: f"Action_{c}_Count" for c in counts.columns if c != "Sold Party"})
    return summary.merge(counts, on="Sold Party", how="left")


def build_party_summary(
    sales_all: pd.DataFrame,
    sales_only: pd.DataFrame,
    bids: pd.DataFrame,
    actions: pd.DataFrame,
    deals: pd.DataFrame,
    overrides: pd.DataFrame,
    removed: pd.DataFrame,
    settings: dict[str, Any],
    as_of: pd.Timestamp,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    recent_action_days = int(float(settings.get("RECENT_ACTION_DAYS", 90)))
    base = sales_all.groupby("Sold Party").agg(
        First_Recorded_Date=("Date", "min"),
        Last_Recorded_Date=("Date", "max"),
        Sales_Person=("Sales_Person", first_nonblank),
        Country=("Country", first_nonblank),
        City=("City", first_nonblank),
    ).reset_index()

    sale_sum = sales_only.groupby("Sold Party").agg(
        First_Sale_Date=("Date", "min"),
        Last_Sale_Date=("Date", "max"),
        Sale_Days=("Date", "nunique"),
        Sale_Stone_Count=("Serial_Key", "size"),
    ).reset_index()
    bid_sum = bids.groupby("Sold Party").agg(
        Last_Bid_Date=("Date", "max"),
        Bid_Count=("Type", "size"),
    ).reset_index() if not bids.empty else pd.DataFrame(columns=["Sold Party", "Last_Bid_Date", "Bid_Count"])
    action_sum = aggregate_actions(actions, as_of, recent_action_days)

    if deals.empty:
        deal_sum = pd.DataFrame(columns=["Sold Party"])
        interval_history = pd.DataFrame(columns=["Sold Party", "Deal_ID", "Deal_Start_Date", "Previous_Deal_Start_Date", "Interval_Days"])
    else:
        interval_history = deals[deals["Interval_Days"].notna()].copy()
        deal_sum = deals.groupby("Sold Party").agg(
            First_Deal_Date=("Deal_Start_Date", "min"),
            Last_Deal_Start_Date=("Deal_Start_Date", "max"),
            Last_Deal_Date=("Deal_End_Date", "max"),
            Deal_Bucket_Count=("Deal_ID", "nunique"),
            Total_Deal_Stones=("Deal_Total_Stones", "sum"),
            Avg_Deal_Stones=("Deal_Total_Stones", "mean"),
            Max_Deal_Stones=("Deal_Total_Stones", "max"),
            Median_Deal_Stones=("Deal_Total_Stones", "median"),
        ).reset_index()
        interval_stats = interval_history.groupby("Sold Party").agg(
            Interval_Count=("Interval_Days", "count"),
            Median_Interval_Days=("Interval_Days", "median"),
            Average_Interval_Days=("Interval_Days", "mean"),
            Min_Interval_Days=("Interval_Days", "min"),
            Max_Interval_Days=("Interval_Days", "max"),
            Interval_P25_Days=("Interval_Days", lambda s: s.quantile(0.25)),
            Interval_P75_Days=("Interval_Days", lambda s: s.quantile(0.75)),
        ).reset_index()
        deal_sum = deal_sum.merge(interval_stats, on="Sold Party", how="left")

    party = base.merge(sale_sum, on="Sold Party", how="outer")
    party = party.merge(bid_sum, on="Sold Party", how="left")
    party = party.merge(action_sum, on="Sold Party", how="left")
    party = party.merge(deal_sum, on="Sold Party", how="left")

    numeric_zero = [
        "Sale_Days", "Sale_Stone_Count", "Bid_Count", "Total_Action_Count", "Recent_Action_Count",
        "Action_Weighted_Score", "Deal_Bucket_Count", "Total_Deal_Stones", "Avg_Deal_Stones",
        "Max_Deal_Stones", "Median_Deal_Stones", "Interval_Count",
    ]
    for c in numeric_zero:
        if c not in party.columns:
            party[c] = 0
        party[c] = pd.to_numeric(party[c], errors="coerce").fillna(0)

    for c in ["Median_Interval_Days", "Average_Interval_Days", "Min_Interval_Days", "Max_Interval_Days", "Interval_P25_Days", "Interval_P75_Days"]:
        if c not in party.columns:
            party[c] = np.nan
        party[c] = pd.to_numeric(party[c], errors="coerce")

    # Normalize overrides.
    ov = overrides.copy()
    if "Sold Party" not in ov.columns:
        ov["Sold Party"] = pd.Series(dtype=str)
    if "Override" not in ov.columns:
        ov["Override"] = "AUTO"
    if "Manual_Interval_Days" not in ov.columns:
        ov["Manual_Interval_Days"] = np.nan
    if "Notes" not in ov.columns:
        ov["Notes"] = ""
    ov["Sold Party"] = ov["Sold Party"].fillna("").astype(str).str.strip()
    ov["Override"] = normalize_upper(ov["Override"], "AUTO")
    ov = ov[ov["Sold Party"].ne("")].drop_duplicates("Sold Party", keep="last")
    party = party.merge(ov[["Sold Party", "Override", "Manual_Interval_Days", "Notes"]], on="Sold Party", how="left")
    party["Override"] = party["Override"].fillna("AUTO").astype(str).str.upper()
    party["Manual_Interval_Days"] = pd.to_numeric(party["Manual_Interval_Days"], errors="coerce")
    party["Notes"] = party["Notes"].fillna("")

    removed_set = set()
    if "Sold Party" in removed.columns:
        removed_set = set(removed["Sold Party"].dropna().astype(str).str.strip())
    party["Removed_Flag"] = party["Sold Party"].isin(removed_set).map({True: "YES", False: "NO"})

    min_buckets = int(float(settings.get("MIN_DEAL_BUCKETS_FOR_MAKER", 3)))
    min_intervals = int(float(settings.get("MIN_INTERVALS_FOR_AUTO", 2)))
    due_soon_ratio = float(settings.get("DUE_SOON_RATIO", 0.8))
    due_now_ratio = float(settings.get("DUE_NOW_RATIO", 1.0))
    overdue_ratio = float(settings.get("OVERDUE_RATIO", 1.3))
    require_action = yes(settings.get("REQUIRE_RECENT_ACTION", "NO"))
    include_due_soon = yes(settings.get("INCLUDE_DUE_SOON", "NO"))

    def maker_class(row: pd.Series) -> str:
        buckets = int(row.get("Deal_Bucket_Count", 0))
        intervals = int(row.get("Interval_Count", 0))
        if buckets >= min_buckets and intervals >= min_intervals:
            return "DEAL MAKER"
        if buckets >= 2 and intervals >= 1:
            return "REPEAT DEAL BUYER - LIMITED HISTORY"
        if buckets == 1:
            return "ONE REAL DEAL"
        return "NO REAL DEAL HISTORY"

    party["Deal_Maker_Class"] = party.apply(maker_class, axis=1)
    party["Effective_Interval_Days"] = party["Manual_Interval_Days"].where(
        party["Manual_Interval_Days"].gt(0), party["Median_Interval_Days"]
    )
    party["Interval_Source"] = np.where(
        party["Manual_Interval_Days"].gt(0), "MANUAL", np.where(party["Median_Interval_Days"].notna(), "MEDIAN HISTORY", "NONE")
    )
    party["Data_As_Of_Date"] = as_of
    party["Days_Since_Last_Deal"] = (as_of - party["Last_Deal_Date"]).dt.days
    party["Days_Since_Last_Sale"] = (as_of - party["Last_Sale_Date"]).dt.days
    party["Days_Since_Last_Action"] = (as_of - party["Last_Action_Date"]).dt.days
    party["Next_Expected_Deal_Date"] = party.apply(
        lambda r: r["Last_Deal_Date"] + pd.Timedelta(days=int(round(r["Effective_Interval_Days"])))
        if pd.notna(r.get("Last_Deal_Date")) and pd.notna(r.get("Effective_Interval_Days")) and r.get("Effective_Interval_Days", 0) > 0
        else pd.NaT,
        axis=1,
    )
    party["Interval_Progress_Ratio"] = party["Days_Since_Last_Deal"] / party["Effective_Interval_Days"]

    def interval_status(row: pd.Series) -> str:
        override = str(row.get("Override", "AUTO")).upper()
        if override == "FORCE_INCLUDE":
            return "MANUAL INCLUDE"
        if override == "FORCE_EXCLUDE" or row.get("Removed_Flag") == "YES":
            return "MANUAL EXCLUDE"
        if row.get("Deal_Maker_Class") != "DEAL MAKER":
            return "INSUFFICIENT REPEAT HISTORY"
        ratio = row.get("Interval_Progress_Ratio")
        if pd.isna(ratio):
            return "NO INTERVAL"
        if ratio < due_soon_ratio:
            return "EARLY"
        if ratio < due_now_ratio:
            return "DUE SOON"
        if ratio <= overdue_ratio:
            return "DUE NOW"
        return "OVERDUE"

    party["Interval_Status"] = party.apply(interval_status, axis=1)
    party["Recent_Action_Flag"] = party["Days_Since_Last_Action"].le(recent_action_days).fillna(False).map({True: "YES", False: "NO"})

    def auto_include(row: pd.Series) -> tuple[str, str]:
        override = str(row.get("Override", "AUTO")).upper()
        if override == "FORCE_EXCLUDE" or row.get("Removed_Flag") == "YES":
            return "NO", "Excluded by user control"
        if override == "FORCE_INCLUDE":
            return "YES", "Forced include by user control"
        allowed_status = {"DUE NOW", "OVERDUE"}
        if include_due_soon:
            allowed_status.add("DUE SOON")
        if row.get("Deal_Maker_Class") != "DEAL MAKER":
            return "NO", "Not enough repeat deal history"
        if row.get("Interval_Status") not in allowed_status:
            return "NO", f"Interval status is {row.get('Interval_Status')}"
        if require_action and row.get("Recent_Action_Flag") != "YES":
            return "NO", f"No action in last {recent_action_days} days"
        return "YES", "Interval gate passed"

    include_results = party.apply(auto_include, axis=1)
    party["Auto_Include"] = [x[0] for x in include_results]
    party["Auto_Include_Reason"] = [x[1] for x in include_results]

    party["Deal_Interval_Remark"] = party.apply(
        lambda r: (
            f"{r['Deal_Maker_Class']} | {int(r['Deal_Bucket_Count'])} deals | "
            f"median interval {round(r['Effective_Interval_Days'], 1) if pd.notna(r['Effective_Interval_Days']) else 'NA'} days | "
            f"last deal {int(r['Days_Since_Last_Deal']) if pd.notna(r['Days_Since_Last_Deal']) else 'NA'} days ago | "
            f"{r['Interval_Status']} | include {r['Auto_Include']}"
        ), axis=1
    )

    party = party.sort_values(
        ["Auto_Include", "Deal_Maker_Class", "Interval_Progress_Ratio", "Deal_Bucket_Count"],
        ascending=[False, True, False, False], na_position="last"
    ).reset_index(drop=True)
    return party, interval_history


def build_matcher_profile(
    sales_only: pd.DataFrame,
    bids: pd.DataFrame,
    actions: pd.DataFrame,
    party_summary: pd.DataFrame,
    settings: dict[str, Any],
    as_of: pd.Timestamp,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    lookback = int(float(settings.get("PROFILE_LOOKBACK_DAYS", 0) or 0))
    min_sales = int(float(settings.get("MIN_PROFILE_SALE_STONES", 1)))
    sale_profile_data = sales_only.copy()
    if lookback > 0:
        sale_profile_data = sale_profile_data[sale_profile_data["Date"] >= as_of - pd.Timedelta(days=lookback)]

    agg = sale_profile_data.groupby(["Sold Party", "Supertag"]).agg(
        Sale_Stone_Count=("Serial_Key", "size"),
        Sale_Days=("Date", "nunique"),
        First_Sale_Date=("Date", "min"),
        Last_Sale_Date=("Date", "max"),
        Avg_Sale_Discount=("Disc%", "mean"),
        Majority_Shade_Score=("Shade_Score", majority_score),
        Max_Shade_Score=("Shade_Score", "max"),
        Majority_Milky_Score=("Milky_Score", majority_score),
        Max_Milky_Score=("Milky_Score", "max"),
        Majority_Open_Score=("Open_Score", majority_score),
        Max_Open_Score=("Open_Score", "max"),
        Majority_Blk_Score=("Blk_Score", majority_score),
        Max_Blk_Score=("Blk_Score", "max"),
        Sales_Person=("Sales_Person", first_nonblank),
        Country=("Country", first_nonblank),
        City=("City", first_nonblank),
    ).reset_index()
    agg = agg[agg["Sale_Stone_Count"] >= min_sales].copy()

    if not bids.empty:
        bid_agg = bids.groupby(["Sold Party", "Supertag"]).agg(
            Bid_Count=("Type", "size"),
            Last_Bid_Date=("Date", "max"),
            Avg_Bid_Discount=("Disc%", "mean"),
        ).reset_index()
        agg = agg.merge(bid_agg, on=["Sold Party", "Supertag"], how="left")
    else:
        agg["Bid_Count"] = 0
        agg["Last_Bid_Date"] = pd.NaT
        agg["Avg_Bid_Discount"] = np.nan

    if not actions.empty:
        actions2 = actions.copy()
        actions2["Action_Weight"] = actions2["Type"].map(ACTION_WEIGHTS).fillna(0.25)
        act_agg = actions2.groupby(["Sold Party", "Supertag"]).agg(
            Action_Count=("Type", "size"),
            Last_Action_Date=("Date", "max"),
            Action_Weighted_Score=("Action_Weight", "sum"),
        ).reset_index()
        agg = agg.merge(act_agg, on=["Sold Party", "Supertag"], how="left")
    else:
        agg["Action_Count"] = 0
        agg["Last_Action_Date"] = pd.NaT
        agg["Action_Weighted_Score"] = 0

    for c in ["Bid_Count", "Action_Count", "Action_Weighted_Score"]:
        agg[c] = pd.to_numeric(agg[c], errors="coerce").fillna(0)

    for dim in ["Shade", "Milky", "Open", "Blk"]:
        agg[f"Majority_{dim}"] = agg[f"Majority_{dim}_Score"].apply(lambda x: score_to_text(x, dim))
        agg[f"Max_{dim}_Accepted"] = agg[f"Max_{dim}_Score"].apply(lambda x: score_to_text(x, dim))

    party_cols = [
        "Sold Party", "Deal_Maker_Class", "Deal_Bucket_Count", "Median_Interval_Days",
        "Effective_Interval_Days", "Last_Deal_Date", "Days_Since_Last_Deal", "Next_Expected_Deal_Date",
        "Interval_Status", "Recent_Action_Flag", "Auto_Include", "Auto_Include_Reason",
        "Override", "Deal_Interval_Remark",
    ]
    agg = agg.merge(party_summary[party_cols], on="Sold Party", how="left", validate="many_to_one")

    agg["Profile_Priority_Score"] = (
        agg["Sale_Stone_Count"] * 10
        + agg["Sale_Days"] * 5
        + agg["Bid_Count"] * 2
        + agg["Action_Weighted_Score"]
    )
    agg["Total_Stones_Bought"] = agg["Sale_Stone_Count"]
    agg["Avg_Discount_Bid"] = agg["Avg_Sale_Discount"]  # legacy compatibility only

    output_cols = [
        "Auto_Include", "Auto_Include_Reason", "Interval_Status", "Deal_Maker_Class",
        "Deal_Bucket_Count", "Effective_Interval_Days", "Last_Deal_Date", "Days_Since_Last_Deal",
        "Next_Expected_Deal_Date", "Recent_Action_Flag", "Override", "Sold Party", "Sales_Person",
        "Country", "City", "Supertag", "Sale_Stone_Count", "Sale_Days", "First_Sale_Date",
        "Last_Sale_Date", "Avg_Sale_Discount", "Bid_Count", "Last_Bid_Date", "Avg_Bid_Discount",
        "Action_Count", "Last_Action_Date", "Action_Weighted_Score", "Profile_Priority_Score",
        "Majority_Shade", "Max_Shade_Accepted", "Majority_Milky", "Max_Milky_Accepted",
        "Majority_Open", "Max_Open_Accepted", "Majority_Blk", "Max_Blk_Accepted",
        "Total_Stones_Bought", "Avg_Discount_Bid", "Deal_Interval_Remark",
    ]
    all_profile = agg[output_cols].sort_values(
        ["Auto_Include", "Sold Party", "Profile_Priority_Score"], ascending=[False, True, False]
    ).reset_index(drop=True)
    eligible = all_profile[all_profile["Auto_Include"] == "YES"].copy().reset_index(drop=True)

    dup = all_profile.duplicated(["Sold Party", "Supertag"], keep=False)
    if dup.any():
        raise ValueError("Matcher profile is not unique by Sold Party + Supertag. This should never occur.")
    return all_profile, eligible


def build_legacy_audit(path: Path | None) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    if not path:
        return pd.DataFrame(columns=["Check", "Value", "Severity", "Comment"])
    if not path.exists():
        return pd.DataFrame([{"Check": "Legacy deal file", "Value": "Not found", "Severity": "WARNING", "Comment": str(path)}])
    try:
        df = pd.read_excel(path)
        rows.append({"Check": "Legacy rows", "Value": len(df), "Severity": "INFO", "Comment": ""})
        if {"Sold Party", "Supertag"}.issubset(df.columns):
            unique = df[["Sold Party", "Supertag"]].drop_duplicates().shape[0]
            dup_rows = len(df) - unique
            dup_keys = int(df.duplicated(["Sold Party", "Supertag"], keep=False).groupby([df["Sold Party"], df["Supertag"]]).any().sum()) if len(df) else 0
            rows.append({"Check": "Unique Party + Supertag", "Value": unique, "Severity": "INFO", "Comment": "Required grain for stock matcher."})
            rows.append({"Check": "Rows above unique grain", "Value": dup_rows, "Severity": "ERROR" if dup_rows else "OK", "Comment": "These rows can create many-to-many stock matches."})
            rows.append({"Check": "Duplicated Party + Supertag keys", "Value": dup_keys, "Severity": "ERROR" if dup_keys else "OK", "Comment": "Must be consolidated before Step 2."})
        for c in ["Max_Shade_Accepted", "Max_Milky_Accepted", "Max_Open_Accepted", "Max_Blk_Accepted"]:
            if c not in df.columns:
                rows.append({"Check": f"Missing {c}", "Value": 1, "Severity": "ERROR", "Comment": "Required quality limit."})
    except Exception as exc:
        rows.append({"Check": "Legacy file read error", "Value": str(exc), "Severity": "ERROR", "Comment": str(path)})
    return pd.DataFrame(rows)


def code_audit_table() -> pd.DataFrame:
    rows = [
        ("CRITICAL", "Step 1 does not filter Type == SALE before sales aggregation.", "BID rows can become completed purchases and deal buckets.", "Filter completed sales first; retain BID separately as intent."),
        ("CRITICAL", "Step 1 counts Supertag rows instead of unique stone serials.", "Duplicate rows can inflate deal size.", "Count unique Serial No per party and business date."),
        ("HIGH", "Average interval is used as the main cadence.", "Large gaps distort the expected interval.", "Use median interval as the primary cadence; show average only as context."),
        ("HIGH", "One large deal can qualify as a strong buyer through average deal stones.", "No repeat interval exists for a one-time buyer.", "Require minimum deal buckets and measured intervals for auto inclusion."),
        ("HIGH", "Old deal-bucket selection is strongest-first across the full party history.", "Overlapping windows can be selected inconsistently.", "Use deterministic chronological dynamic programming."),
        ("HIGH", "Step 1 has no machine-safe Auto_Include output.", "Manual checking is still required.", "Create strict Eligible_Parties and Stock_Fetcher_Input sheets."),
        ("HIGH", "Step 2 does not use interval eligibility before stock matching.", "Deals can be generated for early/not-due parties.", "Read only Stock_Fetcher_Input with Auto_Include == YES."),
        ("CRITICAL", "Step 2 merges a duplicated deal profile on Supertag.", "Many-to-many rows can leave an arbitrary quality profile.", "Validate one row per Sold Party + Supertag before matching."),
        ("HIGH", "Step 2 calculates open-inclusion score but does not gate on it.", "Stones above accepted open level can pass.", "Include open score in limit and tier checks."),
        ("HIGH", "Unknown quality values rank as zero in Step 2.", "Unknown values are treated as best quality.", "Flag unknowns and reject or require review."),
        ("HIGH", "Pair injection can add a partner from full stock after filters.", "Partner may not match profile/exclusion rules.", "Only pair stones that independently pass all rules."),
        ("MEDIUM", "File names and paths are hardcoded.", "Each run requires code editing.", "Use command-line parameters and one reusable control workbook."),
    ]
    return pd.DataFrame(rows, columns=["Severity", "Finding", "Risk", "Replacement Fix"])


def create_or_update_override_rows(existing: pd.DataFrame, parties: Iterable[str]) -> pd.DataFrame:
    required = ["Sold Party", "Override", "Manual_Interval_Days", "Notes"]
    out = existing.copy()
    for c in required:
        if c not in out.columns:
            out[c] = "" if c != "Manual_Interval_Days" else np.nan
    out = out[required]
    existing_parties = set(out["Sold Party"].dropna().astype(str).str.strip())
    new_rows = [
        {"Sold Party": p, "Override": "AUTO", "Manual_Interval_Days": np.nan, "Notes": ""}
        for p in sorted(set(parties)) if p not in existing_parties
    ]
    if new_rows:
        out = pd.concat([out, pd.DataFrame(new_rows)], ignore_index=True)
    out["Override"] = out["Override"].fillna("AUTO").replace("", "AUTO")
    return out.sort_values("Sold Party").reset_index(drop=True)


def dataframe_to_sheet(ws, df: pd.DataFrame, title: str | None = None) -> None:
    if title:
        ws.append([title])
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(1, len(df.columns)))
        ws["A1"].font = Font(bold=True, color="FFFFFF", size=14)
        ws["A1"].fill = PatternFill("solid", fgColor="17365D")
        ws["A1"].alignment = Alignment(horizontal="left")
        header_row = 3
    else:
        header_row = 1
    for col_idx, col in enumerate(df.columns, 1):
        ws.cell(header_row, col_idx, col)
    for r_idx, row in enumerate(df.itertuples(index=False, name=None), header_row + 1):
        for c_idx, value in enumerate(row, 1):
            if pd.isna(value):
                value = None
            elif isinstance(value, pd.Timestamp):
                value = value.to_pydatetime()
            elif isinstance(value, np.generic):
                value = value.item()
            ws.cell(r_idx, c_idx, value)
    style_table(ws, header_row, len(df), len(df.columns))


def style_table(ws, header_row: int, data_rows: int, columns: int) -> None:
    dark = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[header_row]:
        if cell.column <= columns:
            cell.fill = dark
            cell.font = Font(color="FFFFFF", bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = f"A{header_row + 1}"
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(max(1, columns))}{header_row + data_rows}"
    ws.sheet_view.showGridLines = False
    date_names = {
        "Date", "Data_As_Of_Date", "First_Sale_Date", "Last_Sale_Date", "Last_Bid_Date", "Last_Action_Date",
        "First_Deal_Date", "Last_Deal_Start_Date", "Last_Deal_Date", "Deal_Start_Date", "Deal_End_Date",
        "Previous_Deal_Start_Date", "Next_Expected_Deal_Date", "First_Recorded_Date", "Last_Recorded_Date",
    }
    headers = {ws.cell(header_row, c).value: c for c in range(1, columns + 1)}
    for name in date_names.intersection(headers):
        idx = headers[name]
        for r in range(header_row + 1, header_row + data_rows + 1):
            ws.cell(r, idx).number_format = "DD-MM-YYYY"
    for name, idx in headers.items():
        if name and any(token in str(name) for token in ["Discount", "Ratio"]):
            for r in range(header_row + 1, header_row + data_rows + 1):
                ws.cell(r, idx).number_format = "0.00"
    for idx in range(1, columns + 1):
        values = [clean_text(ws.cell(r, idx).value) for r in range(header_row, min(header_row + data_rows + 1, header_row + 101))]
        width = min(max(max((len(v) for v in values), default=8) + 2, 10), 45)
        ws.column_dimensions[get_column_letter(idx)].width = width
    for row in ws.iter_rows(min_row=header_row + 1, max_row=header_row + data_rows, max_col=columns):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=False)


def build_dashboard(ws, party: pd.DataFrame, deals: pd.DataFrame, eligible: pd.DataFrame, as_of: pd.Timestamp) -> None:
    ws.sheet_view.showGridLines = False
    ws["A1"] = "DEALMAKER INTERVAL CONTROL"
    ws["A1"].font = Font(size=20, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="17365D")
    ws.merge_cells("A1:N2")
    ws["A3"] = f"Version {VERSION} | Data as of {as_of.strftime('%d-%m-%Y')}"
    ws["A3"].font = Font(italic=True, color="666666")

    kpis = [
        ("Parties analyzed", len(party)),
        ("Deal makers", int((party["Deal_Maker_Class"] == "DEAL MAKER").sum()) if not party.empty else 0),
        ("Real deal buckets", len(deals)),
        ("Interval eligible parties", int((party["Auto_Include"] == "YES").sum()) if not party.empty else 0),
        ("Stock input profile rows", len(eligible)),
    ]
    for i, (label, value) in enumerate(kpis):
        row_block = i // 3
        col_block = i % 3
        col = 1 + col_block * 2
        row = 5 + row_block * 3
        ws.cell(row, col, label)
        ws.cell(row + 1, col, value)
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + 1)
        ws.merge_cells(start_row=row + 1, start_column=col, end_row=row + 2, end_column=col + 1)
        ws.cell(row, col).fill = PatternFill("solid", fgColor="D9EAF7")
        ws.cell(row, col).font = Font(bold=True)
        ws.cell(row, col).alignment = Alignment(horizontal="center")
        ws.cell(row + 1, col).fill = PatternFill("solid", fgColor="E2F0D9")
        ws.cell(row + 1, col).font = Font(size=18, bold=True)
        ws.cell(row + 1, col).alignment = Alignment(horizontal="center", vertical="center")

    top_cols = [
        "Sold Party", "Deal_Maker_Class", "Deal_Bucket_Count", "Median_Interval_Days",
        "Days_Since_Last_Deal", "Interval_Status", "Auto_Include",
    ]
    top = party[top_cols].head(15) if not party.empty else pd.DataFrame(columns=top_cols)
    start = 12
    ws.cell(start, 1, "Top party interval summary")
    ws.cell(start, 1).font = Font(bold=True, size=13, color="FFFFFF")
    ws.cell(start, 1).fill = PatternFill("solid", fgColor="1F4E78")
    ws.merge_cells(start_row=start, start_column=1, end_row=start, end_column=len(top_cols))
    for c, name in enumerate(top.columns, 1):
        ws.cell(start + 1, c, name)
    for r, row in enumerate(top.itertuples(index=False, name=None), start + 2):
        for c, value in enumerate(row, 1):
            if pd.isna(value):
                value = None
            ws.cell(r, c, value)
    style_table(ws, start + 1, len(top), len(top_cols))

    if not top.empty:
        chart = BarChart()
        chart.title = "Deal buckets by party"
        chart.y_axis.title = "Deal buckets"
        chart.x_axis.title = "Party"
        data = Reference(ws, min_col=3, min_row=start + 1, max_row=start + 1 + len(top))
        cats = Reference(ws, min_col=1, min_row=start + 2, max_row=start + 1 + len(top))
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.height = 7
        chart.width = 13
        ws.add_chart(chart, "I12")

    widths = {"A": 38, "B": 29, "C": 15, "D": 18, "E": 18, "F": 16, "G": 13}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for col in range(8, 15):
        ws.column_dimensions[get_column_letter(col)].width = 12
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = "A1:N28"


def write_control_workbook(
    output: Path,
    settings: dict[str, Any],
    overrides: pd.DataFrame,
    removed: pd.DataFrame,
    exclusions: pd.DataFrame,
    party: pd.DataFrame,
    deals: pd.DataFrame,
    intervals: pd.DataFrame,
    all_profile: pd.DataFrame,
    eligible_profile: pd.DataFrame,
    quality: pd.DataFrame,
    legacy_audit: pd.DataFrame,
    as_of: pd.Timestamp,
) -> None:
    wb = Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet("Dashboard")
    build_dashboard(ws, party, deals, eligible_profile, as_of)

    summary_rows = [
        ("Engine version", VERSION),
        ("Data as-of date", as_of),
        ("Parties analyzed", len(party)),
        ("Completed SALE rows used", int(quality.loc[(quality["Dataset"] == "Sales") & (quality["Check"] == "Completed SALE rows after dedupe"), "Value"].sum()) if not quality.empty else 0),
        ("Real deal buckets", len(deals)),
        ("Deal makers", int((party["Deal_Maker_Class"] == "DEAL MAKER").sum()) if not party.empty else 0),
        ("Eligible parties", int((party["Auto_Include"] == "YES").sum()) if not party.empty else 0),
        ("Stock fetcher profile rows", len(eligible_profile)),
        ("Primary interval measure", "Median start-to-start interval"),
        ("Automatic output rule", "Repeat deal maker + due interval + optional recent-action gate"),
    ]
    summary = pd.DataFrame(summary_rows, columns=["Metric", "Value"])
    dataframe_to_sheet(wb.create_sheet("Run_Summary"), summary)
    dataframe_to_sheet(wb.create_sheet("Party_Deal_Summary"), party)
    dataframe_to_sheet(wb.create_sheet("Deal_Buckets"), deals)
    interval_cols = [
        "Sold Party", "Deal_ID", "Party_Deal_Number", "Deal_Start_Date", "Deal_End_Date",
        "Previous_Deal_Start_Date", "Interval_Days", "Deal_Total_Stones", "Deal_Days",
        "Sales_Person", "Country", "City",
    ]
    dataframe_to_sheet(wb.create_sheet("Deal_Interval_History"), intervals[[c for c in interval_cols if c in intervals.columns]])
    eligible_parties = party[party["Auto_Include"] == "YES"].copy()
    dataframe_to_sheet(wb.create_sheet("Eligible_Parties"), eligible_parties)
    dataframe_to_sheet(wb.create_sheet("Stock_Fetcher_Input"), eligible_profile)
    dataframe_to_sheet(wb.create_sheet("Matcher_Profile_All"), all_profile)
    dataframe_to_sheet(wb.create_sheet("Data_Quality"), quality)
    dataframe_to_sheet(wb.create_sheet("Code_Audit"), code_audit_table())
    dataframe_to_sheet(wb.create_sheet("Legacy_Input_Audit"), legacy_audit)

    # Input sheets.
    settings_df = pd.DataFrame(DEFAULT_SETTINGS, columns=["Setting", "Default", "Description"])
    settings_df["Value"] = settings_df["Setting"].map(settings)
    settings_df = settings_df[["Setting", "Value", "Default", "Description"]]
    ws_settings = wb.create_sheet("Settings")
    dataframe_to_sheet(ws_settings, settings_df, "USER INPUT - EDIT ONLY THE BLUE VALUE CELLS")
    for r in range(4, 4 + len(settings_df)):
        ws_settings.cell(r, 2).font = Font(color="0000FF")
        ws_settings.cell(r, 2).fill = PatternFill("solid", fgColor="FFF2CC")
        ws_settings.cell(r, 2).protection = Protection(locked=False)
    ws_settings.protection.sheet = True
    ws_settings.protection.enable()

    ws_ov = wb.create_sheet("Party_Overrides")
    dataframe_to_sheet(ws_ov, overrides, "USER INPUT - AUTO / FORCE_INCLUDE / FORCE_EXCLUDE")
    if len(overrides) > 0:
        dv = DataValidation(type="list", formula1='"AUTO,FORCE_INCLUDE,FORCE_EXCLUDE"', allow_blank=False)
        ws_ov.add_data_validation(dv)
        dv.add(f"B4:B{3 + len(overrides)}")
        for r in range(4, 4 + len(overrides)):
            for c in range(2, 5):
                ws_ov.cell(r, c).font = Font(color="0000FF")
                ws_ov.cell(r, c).fill = PatternFill("solid", fgColor="FFF2CC")

    ws_removed = wb.create_sheet("Removed_Parties")
    dataframe_to_sheet(ws_removed, removed, "USER INPUT - PARTIES NEVER TO FETCH")
    ws_ex = wb.create_sheet("Exclusion_Rules")
    dataframe_to_sheet(ws_ex, exclusions, "USER INPUT - OPTIONAL STOCK EXCLUSION RULES")

    guide = pd.DataFrame([
        (1, "Place full Sales and Actions files in one folder."),
        (2, "Run Step 1. It updates this workbook and preserves Settings/Overrides/Exclusions."),
        (3, "Review Party_Deal_Summary. No manual filtering is required for normal runs."),
        (4, "Stock_Fetcher_Input contains only interval-eligible Party + Supertag profiles."),
        (5, "Run Step 2 using this workbook and the current stock file."),
        (6, "Use FORCE_INCLUDE only for a business exception; use FORCE_EXCLUDE or Removed_Parties to suppress."),
        (7, "Do not copy the legacy raw deal file into Step 2; use Stock_Fetcher_Input from this workbook."),
    ], columns=["Step", "Instruction"])
    dataframe_to_sheet(wb.create_sheet("User_Guide"), guide)

    # Conditional formatting on include/status.
    for sheet_name in ["Party_Deal_Summary", "Eligible_Parties", "Stock_Fetcher_Input", "Matcher_Profile_All"]:
        ws2 = wb[sheet_name]
        headers = {ws2.cell(1, c).value: c for c in range(1, ws2.max_column + 1)}
        if "Auto_Include" in headers and ws2.max_row >= 2:
            col_letter = get_column_letter(headers["Auto_Include"])
            green = PatternFill("solid", fgColor="C6EFCE")
            red = PatternFill("solid", fgColor="FFC7CE")
            ws2.conditional_formatting.add(
                f"A2:{get_column_letter(ws2.max_column)}{ws2.max_row}",
                FormulaRule(formula=[f'${col_letter}2="YES"'], fill=green),
            )
            ws2.conditional_formatting.add(
                f"A2:{get_column_letter(ws2.max_column)}{ws2.max_row}",
                FormulaRule(formula=[f'${col_letter}2="NO"'], fill=red),
            )

    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.save(output)


def run(args: argparse.Namespace) -> None:
    sales_path = Path(args.sales).resolve()
    actions_path = Path(args.actions).resolve()
    control_path = Path(args.control).resolve()
    legacy_path = Path(args.legacy_deal).resolve() if args.legacy_deal else None

    settings, overrides, removed, exclusions = load_control_inputs(control_path)
    quality_log: list[dict[str, Any]] = []
    sales_mapping = load_manual_mapping(control_path, "SALES")
    actions_mapping = load_manual_mapping(control_path, "ACTIONS")
    sales = prepare_dataset(
        sales_path, "Sales", quality_log, args.sales_sheet or None, args.sales_header_row, sales_mapping
    )
    actions = prepare_dataset(
        actions_path, "Actions", quality_log, args.actions_sheet or None, args.actions_header_row, actions_mapping
    )
    sales = dedupe_business_rows(sales, "Sales", quality_log)
    exact_action_duplicates = int(actions.duplicated(keep=False).sum())
    quality_log.append({
        "Dataset": "Actions",
        "Check": "Exact duplicate-looking action rows retained",
        "Value": exact_action_duplicates,
        "Severity": "WARNING" if exact_action_duplicates else "OK",
    })

    sale_types = parse_csv_set(settings.get("SALE_TYPE_VALUES", "SALE"))
    bid_types = parse_csv_set(settings.get("BID_TYPE_VALUES", "BID"))
    sales_only = sales[sales["Type"].isin(sale_types)].copy()
    bids = sales[sales["Type"].isin(bid_types)].copy()
    unknown_sales_types = sorted(set(sales["Type"]) - sale_types - bid_types)

    quality_log.extend([
        {"Dataset": "Sales", "Check": "Completed SALE rows after dedupe", "Value": len(sales_only), "Severity": "INFO"},
        {"Dataset": "Sales", "Check": "BID rows retained as intent only", "Value": len(bids), "Severity": "INFO"},
        {"Dataset": "Sales", "Check": "Other Type values", "Value": ", ".join(unknown_sales_types) if unknown_sales_types else "None", "Severity": "WARNING" if unknown_sales_types else "OK"},
    ])
    if sales_only.empty:
        raise ValueError(f"No completed sales found using SALE_TYPE_VALUES={sorted(sale_types)}")

    as_of = determine_as_of(settings, sales, actions)
    min_stones = int(float(settings.get("MIN_DEAL_STONES", 6)))
    max_days = int(float(settings.get("MAX_DEAL_DAYS", 2)))
    deals, _ = build_deal_buckets(sales_only, min_stones, max_days)
    overrides = create_or_update_override_rows(overrides, sorted(set(sales["Sold Party"]) | set(actions["Sold Party"])))
    party, intervals = build_party_summary(sales, sales_only, bids, actions, deals, overrides, removed, settings, as_of)
    all_profile, eligible_profile = build_matcher_profile(sales_only, bids, actions, party, settings, as_of)

    duplicates = int(all_profile.duplicated(["Sold Party", "Supertag"]).sum()) if not all_profile.empty else 0
    quality_log.extend([
        {"Dataset": "Output", "Check": "Real deal buckets", "Value": len(deals), "Severity": "INFO"},
        {"Dataset": "Output", "Check": "Deal makers", "Value": int((party["Deal_Maker_Class"] == "DEAL MAKER").sum()), "Severity": "INFO"},
        {"Dataset": "Output", "Check": "Interval eligible parties", "Value": int((party["Auto_Include"] == "YES").sum()), "Severity": "INFO"},
        {"Dataset": "Output", "Check": "Stock_Fetcher_Input rows", "Value": len(eligible_profile), "Severity": "INFO"},
        {"Dataset": "Output", "Check": "Duplicate Party + Supertag profile rows", "Value": duplicates, "Severity": "ERROR" if duplicates else "OK"},
    ])
    quality = pd.DataFrame(quality_log)
    legacy_audit = build_legacy_audit(legacy_path)

    write_control_workbook(
        control_path, settings, overrides, removed, exclusions, party, deals, intervals,
        all_profile, eligible_profile, quality, legacy_audit, as_of,
    )

    print(f"Success: {control_path}")
    print(f"Data as of: {as_of.strftime('%d-%m-%Y')}")
    print(f"Parties: {len(party)} | deal makers: {(party['Deal_Maker_Class'] == 'DEAL MAKER').sum()} | eligible: {(party['Auto_Include'] == 'YES').sum()}")
    print(f"Real deal buckets: {len(deals)} | Stock_Fetcher_Input rows: {len(eligible_profile)}")


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Build deal-maker interval summary and stock-fetcher input workbook.")
    p.add_argument("--sales", required=True, help="Sales/BID Excel file")
    p.add_argument("--actions", required=True, help="Actions Excel file")
    p.add_argument("--control", default="Dealmaker_Interval_Control.xlsx", help="Reusable control/output workbook")
    p.add_argument("--sales-sheet", default="", help="Optional exact Sales worksheet name; blank auto-detects")
    p.add_argument("--actions-sheet", default="", help="Optional exact Actions worksheet name; blank auto-detects")
    p.add_argument("--sales-header-row", type=int, default=None, help="Optional zero-based Sales header row")
    p.add_argument("--actions-header-row", type=int, default=None, help="Optional zero-based Actions header row")
    p.add_argument("--legacy-deal", default="", help="Optional legacy deal profile file to audit")
    return p


if __name__ == "__main__":
    run(build_parser().parse_args())
