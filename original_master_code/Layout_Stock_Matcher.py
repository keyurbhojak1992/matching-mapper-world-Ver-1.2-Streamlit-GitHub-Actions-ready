"""
Google Colab / Google Drive version
===================================

Project folder:
    /content/drive/MyDrive/Layout

Place these three input workbooks in that folder:
    - one Layout Stock .xlsx file
    - one Sales .xlsx file
    - one Actions .xlsx file

The script auto-detects the files by filename. You can also pass exact names.
Outputs are written to:
    /content/drive/MyDrive/Layout/Output

Matching rules
--------------
Exact keys: Shape + Carat group/rank + Lab
Tolerance keys: Color + Clarity + Fluorescence

TOLERANCE = 0:
    Color, Clarity and Flo must all be exact.

TOLERANCE = 1:
    Color, Clarity and Flo may each differ by at most one rank.

Output row colors
-----------------
Light green : Exact (Color, Clarity and Flo all exact)
Light yellow: Nearby (exactly one of Color/Clarity/Flo differs by one rank)
White       : Multi Nearby (two or three differ by one rank)

One workbook is created per salesperson. Each workbook contains:
    - Navigation sheet
    - one sheet per suggested party

All salesperson workbooks are combined into one ZIP file.

Colab cells:
    from google.colab import drive
    drive.mount('/content/drive')
    !pip -q install pandas openpyxl
    !python "/content/drive/MyDrive/Layout/layout_party_suggestions_colab.py"

Optional exact filenames:
    !python "/content/drive/MyDrive/Layout/layout_party_suggestions_colab.py" \
      --layout "Layout Stock sample file.xlsx" \
      --sales "sales sample file.xlsx" \
      --actions "Actions sample file.xlsx" \
      --tolerance 1
"""

from __future__ import annotations

import argparse
import math
import re
import shutil
import sys
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# -----------------------------------------------------------------------------
# Google Colab / Google Drive defaults
# -----------------------------------------------------------------------------

DEFAULT_PROJECT_DIR = Path("/content/drive/MyDrive/Layout")
DEFAULT_OUTPUT_FOLDER = "Output"
DEFAULT_OUTPUT_ZIP = "Salesperson_Layout_Suggestions.zip"

# -----------------------------------------------------------------------------
# Business configuration
# -----------------------------------------------------------------------------

ACTION_WEIGHTS = {
    "SALE": 100,
    "BID": 75,
    "OFFER": 60,
    "CART": 45,
    "WISHLIST": 35,
    "VIEW": 25,
    "DETAILS": 15,
    "DETAIS": 15,
    "DETAIL": 15,
}

COLOR_ORDER = [
    "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N",
    "O-P", "Q-R", "S-T", "U-V", "W-X", "Y-Z",
]

CLARITY_ORDER = [
    "FL", "IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2",
    "SI3", "I1", "I2", "I3", "I4", "I5",
]

FLO_RANK = {
    "NON": 1,
    "FNT": 2,
    "VSL": 2,
    "SL": 2,
    "MED": 3,
    "STG": 4,
    "VST": 5,
}

COLOR_RANK = {value: index + 1 for index, value in enumerate(COLOR_ORDER)}
CLARITY_RANK = {value: index + 1 for index, value in enumerate(CLARITY_ORDER)}

# Individual colors that belong to grouped color grades.
COLOR_ALIASES = {
    "O": "O-P", "P": "O-P",
    "Q": "Q-R", "R": "Q-R",
    "S": "S-T", "T": "S-T",
    "U": "U-V", "V": "U-V",
    "W": "W-X", "X": "W-X",
    "Y": "Y-Z", "Z": "Y-Z",
}

FLO_ALIASES = {
    "NONE": "NON", "NO": "NON", "NIL": "NON", "N": "NON",
    "FAINT": "FNT",
    "VERY SLIGHT": "VSL", "VERYSLIGHT": "VSL",
    "SLIGHT": "SL",
    "MEDIUM": "MED",
    "STRONG": "STG",
    "VERY STRONG": "VST", "VERYSTRONG": "VST",
}

# Keep empty unless business users approve shape aliases.
SHAPE_ALIASES: dict[str, str] = {}

# Excel colors
HEADER_FILL = "D9E1F2"
HEADER_FONT = "0000FF"
NAV_FILL = "C6EFCE"
EXACT_FILL = "E2F0D9"
NEARBY_FILL = "FFF2CC"
WHITE_FILL = "FFFFFF"
BORDER_COLOR = "7F7F7F"


@dataclass(frozen=True)
class Config:
    tolerance: int = 1
    min_layout_pct: float = 0.10
    min_absolute_stones: int = 5
    min_action_weight: int = 45
    include_match_columns: bool = True
    create_empty_salesperson_files: bool = True


# -----------------------------------------------------------------------------
# Input helpers
# -----------------------------------------------------------------------------

LAYOUT_REQUIRED = {
    "Layout_ID": ["Layout_ID", "Layout ID", "LayoutId"],
    "Segment": ["Segment"],
    "SerialNo": ["SerialNo", "Serial No", "Serial_No"],
    "Shape": ["Shape", "Shp"],
    "Carat": ["Carat"],
    "CaratKey": ["Carat Rank", "Carat_grp", "Carat Group", "CaratGroup"],
    "Color": ["Color", "Col"],
    "Clarity": ["Clarity", "Qua"],
    "Flo": ["Flo", "Fluorescence"],
    "Lab": ["Lab"],
}

HISTORY_REQUIRED = {
    "HistorySerialNo": ["Serial No", "SerialNo", "Serial_No"],
    "Shape": ["Shp", "Shape"],
    "Carat": ["Carat"],
    "CaratKey": ["Carat Rank", "Carat_grp", "Carat Group", "CaratGroup"],
    "Color": ["Col", "Color"],
    "Clarity": ["Qua", "Clarity"],
    "Flo": ["Flo", "Fluorescence"],
    "Lab": ["Lab"],
    "SoldParty": ["Sold Party", "SoldParty", "Party Name"],
    "Country": ["Country"],
    "City": ["City"],
    "Salesperson": ["SALE_PERSON", "Salesperson", "Sales Person"],
    "Type": ["Type", "Action Type"],
    "Date": ["Date", "Action Date", "Sale Date"],
}


def read_first_nonempty_sheet(path: Path) -> pd.DataFrame:
    sheets = pd.read_excel(path, sheet_name=None)
    for _, frame in sheets.items():
        frame = frame.dropna(how="all")
        if not frame.empty:
            frame.columns = [str(c).strip() for c in frame.columns]
            return frame.reset_index(drop=True)
    raise ValueError(f"No non-empty worksheet found in: {path}")


def resolve_column(df: pd.DataFrame, aliases: Iterable[str], required: bool = True) -> Optional[str]:
    exact = {str(c).strip(): c for c in df.columns}
    folded = {str(c).strip().casefold(): c for c in df.columns}
    for alias in aliases:
        if alias in exact:
            return exact[alias]
        if alias.casefold() in folded:
            return folded[alias.casefold()]
    if required:
        raise KeyError(f"Missing required column. Expected one of: {list(aliases)}")
    return None


def standardize_layout(
    df: pd.DataFrame,
) -> tuple[pd.DataFrame, list[tuple[str, str]]]:
    """
    Standardize required columns for matching while preserving the exact original
    Layout Stock headers for output.

    Returns:
        standardized dataframe
        output specification: [(original_header, internal_column_name), ...]
    """
    original_columns = list(df.columns)
    rename_map: dict[str, str] = {}
    for canonical, aliases in LAYOUT_REQUIRED.items():
        source = resolve_column(df, aliases, required=(canonical != "CaratKey"))
        if source is not None:
            rename_map[source] = canonical

    output_spec = [(str(original), rename_map.get(original, original)) for original in original_columns]
    out = df.rename(columns=rename_map).copy()

    # Rebuild CaratKey from numeric Carat so old/pre-existing carat groups
    # cannot mix with the revised standard ranges. If Carat is unavailable,
    # keep any supplied CaratKey as a fallback.
    if "CaratKey" not in out.columns:
        out["CaratKey"] = pd.NA

    supplied_carat_key = out["CaratKey"].copy()
    numeric_carat = pd.to_numeric(out["Carat"], errors="coerce")
    derived_carat_key = numeric_carat.map(derive_carat_group)
    out["CaratKey"] = derived_carat_key.where(numeric_carat.notna(), supplied_carat_key)

    return out, output_spec


def standardize_history(df: pd.DataFrame, source_name: str) -> pd.DataFrame:
    rename_map: dict[str, str] = {}
    optional = {"HistorySerialNo", "Country", "City", "Date", "CaratKey"}
    for canonical, aliases in HISTORY_REQUIRED.items():
        source = resolve_column(df, aliases, required=(canonical not in optional))
        if source is not None:
            rename_map[source] = canonical
    out = df.rename(columns=rename_map).copy()

    # Rebuild CaratKey from numeric Carat so Layout, Sales and Actions all use
    # exactly the same revised standard carat ranges.
    if "CaratKey" not in out.columns:
        out["CaratKey"] = pd.NA

    supplied_carat_key = out["CaratKey"].copy()
    numeric_carat = pd.to_numeric(out["Carat"], errors="coerce")
    derived_carat_key = numeric_carat.map(derive_carat_group)
    out["CaratKey"] = derived_carat_key.where(numeric_carat.notna(), supplied_carat_key)

    for col in ["HistorySerialNo", "Country", "City", "Date"]:
        if col not in out.columns:
            out[col] = pd.NA
    out["SourceFileType"] = source_name
    return out


# -----------------------------------------------------------------------------
# Normalization and ranking
# -----------------------------------------------------------------------------


def clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip()).upper()


def normalize_shape(value: object) -> str:
    shape = clean_text(value)
    return SHAPE_ALIASES.get(shape, shape)


def normalize_carat_key(value: object) -> str:
    text = clean_text(value)
    text = text.replace(" ", "")
    text = text.replace("–", "-").replace("—", "-")
    return text


def normalize_color(value: object) -> str:
    color = clean_text(value).replace(" ", "")
    return COLOR_ALIASES.get(color, color)


def normalize_clarity(value: object) -> str:
    return clean_text(value).replace(" ", "")


def normalize_flo(value: object) -> str:
    flo = clean_text(value)
    return FLO_ALIASES.get(flo, flo)


def normalize_lab(value: object) -> str:
    return clean_text(value).replace(" ", "")


def derive_carat_group(value: object) -> str:
    """
    Return the revised standard carat group used consistently by Layout,
    Sales and Actions data.
    """
    try:
        carat = float(value)
    except (TypeError, ValueError):
        return ""

    bins = [
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
    ]

    if carat >= 15.00:
        return "15+"

    for low, high, label in bins:
        if low <= carat <= high:
            return label

    return ""


def add_normalized_columns(df: pd.DataFrame, is_layout: bool) -> pd.DataFrame:
    out = df.copy()
    out["_Shape"] = out["Shape"].map(normalize_shape)
    out["_CaratKey"] = out["CaratKey"].map(normalize_carat_key)
    out["_Lab"] = out["Lab"].map(normalize_lab)
    out["_Color"] = out["Color"].map(normalize_color)
    out["_Clarity"] = out["Clarity"].map(normalize_clarity)
    out["_Flo"] = out["Flo"].map(normalize_flo)
    out["_ColorRank"] = out["_Color"].map(COLOR_RANK)
    out["_ClarityRank"] = out["_Clarity"].map(CLARITY_RANK)
    out["_FloRank"] = out["_Flo"].map(FLO_RANK)
    if is_layout:
        out["_LayoutSerial"] = out["SerialNo"].astype("string")
    else:
        out["_ActionType"] = out["Type"].map(clean_text)
        out["_ActionWeight"] = out["_ActionType"].map(ACTION_WEIGHTS).fillna(0).astype(int)
        out["_MatchDate"] = pd.to_datetime(out["Date"], errors="coerce")
    return out


# -----------------------------------------------------------------------------
# Matching and recommendation logic
# -----------------------------------------------------------------------------


def create_matches(layout: pd.DataFrame, history: pd.DataFrame, config: Config) -> pd.DataFrame:
    if config.tolerance not in (0, 1):
        raise ValueError("Tolerance must be 0 or 1.")

    layout_n = add_normalized_columns(layout, is_layout=True)
    history_n = add_normalized_columns(history, is_layout=False)

    # Reject records without usable rank values before merging.
    layout_n = layout_n.dropna(subset=["_ColorRank", "_ClarityRank", "_FloRank"])
    history_n = history_n.dropna(subset=["_ColorRank", "_ClarityRank", "_FloRank"])

    exact_keys = ["_Shape", "_CaratKey", "_Lab"]
    merged = history_n.merge(
        layout_n,
        on=exact_keys,
        how="inner",
        suffixes=("_Hist", "_Layout"),
    )

    if merged.empty:
        return merged

    merged["Color Diff"] = (merged["_ColorRank_Hist"] - merged["_ColorRank_Layout"]).abs().astype(int)
    merged["Clarity Diff"] = (merged["_ClarityRank_Hist"] - merged["_ClarityRank_Layout"]).abs().astype(int)
    merged["Flo Diff"] = (merged["_FloRank_Hist"] - merged["_FloRank_Layout"]).abs().astype(int)

    merged = merged[
        (merged["Color Diff"] <= config.tolerance)
        & (merged["Clarity Diff"] <= config.tolerance)
        & (merged["Flo Diff"] <= config.tolerance)
    ].copy()

    if merged.empty:
        return merged

    diff_cols = ["Color Diff", "Clarity Diff", "Flo Diff"]
    merged["_ChangedVariables"] = (merged[diff_cols] > 0).sum(axis=1)
    merged["_TotalDiff"] = merged[diff_cols].sum(axis=1)
    merged["Match Quality"] = merged["_ChangedVariables"].map(
        {0: "Exact", 1: "Nearby", 2: "Multi Nearby", 3: "Multi Nearby"}
    )

    group_keys = ["Salesperson", "SoldParty", "Layout_ID", "_LayoutSerial"]

    # Row color/differences use the closest accepted match for that layout stone.
    best_quality = merged.sort_values(
        by=group_keys + ["_ChangedVariables", "_TotalDiff", "_ActionWeight", "_MatchDate"],
        ascending=[True, True, True, True, True, True, False, False],
        na_position="last",
    ).drop_duplicates(subset=group_keys, keep="first").copy()

    # Buying intent must not be lost when the closest-quality row has a weaker
    # action. Keep the strongest accepted action independently, as Power Query's
    # Max Action Weight logic does.
    strongest = merged.sort_values(
        by=group_keys + ["_ActionWeight", "_MatchDate", "_ChangedVariables", "_TotalDiff"],
        ascending=[True, True, True, True, False, False, True, True],
        na_position="last",
    ).drop_duplicates(subset=group_keys, keep="first")

    strongest = strongest[
        group_keys + ["_ActionWeight", "_ActionType", "HistorySerialNo"]
    ].rename(
        columns={
            "_ActionWeight": "_StrongestActionWeight",
            "_ActionType": "_StrongestActionType",
            "HistorySerialNo": "_StrongestHistorySerialNo",
        }
    )

    latest = (
        merged.groupby(group_keys, dropna=False)["_MatchDate"]
        .max()
        .rename("_LatestMatchDate")
        .reset_index()
    )

    best = best_quality.merge(strongest, on=group_keys, how="left").merge(
        latest, on=group_keys, how="left"
    )
    best["_ActionWeight"] = best["_StrongestActionWeight"]
    best["_ActionType"] = best["_StrongestActionType"]
    best["HistorySerialNo"] = best["_StrongestHistorySerialNo"]
    best["_MatchDate"] = best["_LatestMatchDate"]
    best = best.drop(
        columns=[
            "_StrongestActionWeight", "_StrongestActionType",
            "_StrongestHistorySerialNo", "_LatestMatchDate",
        ]
    )
    return best


def build_recommendations(
    layout: pd.DataFrame,
    best_matches: pd.DataFrame,
    config: Config,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    if best_matches.empty:
        return pd.DataFrame(), pd.DataFrame()

    layout_counts = (
        layout.groupby("Layout_ID", dropna=False)["SerialNo"]
        .nunique()
        .rename("Layout Stone Count")
        .reset_index()
    )

    def first_nonblank(series: pd.Series) -> object:
        for value in series:
            if pd.notna(value) and str(value).strip():
                return value
        return pd.NA

    # Recommendation grain is salesperson + party + Layout_ID. Country and City
    # are descriptive attributes and must not split the distinct-stone count.
    grouped = best_matches.groupby(
        ["Salesperson", "SoldParty", "Layout_ID"],
        dropna=False,
    )

    summary = grouped.agg(
        **{
            "Country": ("Country", first_nonblank),
            "City": ("City", first_nonblank),
            "Matched Stone Count": ("_LayoutSerial", "nunique"),
            "Strongest Action Weight": ("_ActionWeight", "max"),
            "Latest Match Date": ("_MatchDate", "max"),
            "Exact Count": ("Match Quality", lambda s: int((s == "Exact").sum())),
            "Nearby Count": ("Match Quality", lambda s: int((s == "Nearby").sum())),
            "Multi Nearby Count": ("Match Quality", lambda s: int((s == "Multi Nearby").sum())),
        }
    ).reset_index()

    summary = summary.merge(layout_counts, on="Layout_ID", how="left")
    summary["Minimum Required Stones"] = summary["Layout Stone Count"].map(
        lambda n: max(config.min_absolute_stones, math.ceil(float(n) * config.min_layout_pct))
    )
    summary["Preference %"] = (
        summary["Matched Stone Count"] / summary["Layout Stone Count"]
    )
    summary["Recommendation Score"] = (
        summary["Preference %"] * 70
        + (summary["Strongest Action Weight"] / 100) * 30
    )
    summary["Suggested"] = (
        (summary["Matched Stone Count"] >= summary["Minimum Required Stones"])
        & (summary["Strongest Action Weight"] >= config.min_action_weight)
    )

    suggested = summary[summary["Suggested"]].copy()
    if suggested.empty:
        return summary, pd.DataFrame()

    keys = suggested[["Salesperson", "SoldParty", "Layout_ID"]].drop_duplicates()
    suggested_rows = best_matches.merge(
        keys,
        on=["Salesperson", "SoldParty", "Layout_ID"],
        how="inner",
    )
    suggested_rows = suggested_rows.merge(
        suggested[
            [
                "Salesperson", "SoldParty", "Layout_ID",
                "Matched Stone Count", "Layout Stone Count",
                "Minimum Required Stones", "Preference %",
                "Recommendation Score", "Strongest Action Weight",
            ]
        ],
        on=["Salesperson", "SoldParty", "Layout_ID"],
        how="left",
    )
    return summary, suggested_rows


# -----------------------------------------------------------------------------
# Excel output
# -----------------------------------------------------------------------------


def safe_filename(value: object, fallback: str = "Unknown") -> str:
    text = re.sub(r"\s+", " ", str(value or "").strip()) or fallback
    text = re.sub(r"[<>:\"/\\|?*]", "_", text)
    text = text.rstrip(". ")
    return text[:120] or fallback


def salesperson_display_name(value: object) -> str:
    text = re.sub(r"\s+", " ", str(value or "Unknown Salesperson").strip())
    text = re.sub(r"\bMr\.\s*", "Mr ", text, flags=re.IGNORECASE)
    text = re.sub(r"\bMrs\.\s*", "Mrs ", text, flags=re.IGNORECASE)
    text = re.sub(r"\bMs\.\s*", "Ms ", text, flags=re.IGNORECASE)
    return text


def safe_sheet_name(value: object, used: set[str]) -> str:
    base = re.sub(r"[\\/*?:\[\]]", " ", str(value or "Party")).strip()
    base = re.sub(r"\s+", " ", base) or "Party"
    base = base[:31]
    candidate = base
    number = 2
    while candidate.casefold() in used:
        suffix = f" ({number})"
        candidate = f"{base[:31-len(suffix)]}{suffix}"
        number += 1
    used.add(candidate.casefold())
    return candidate


def apply_header_style(ws, row: int, max_col: int) -> None:
    thin = Side(style="thin", color=BORDER_COLOR)
    for cell in ws.iter_cols(min_row=row, max_row=row, min_col=1, max_col=max_col):
        c = cell[0]
        c.fill = PatternFill("solid", fgColor=HEADER_FILL)
        c.font = Font(bold=True, color=HEADER_FONT)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = Border(left=thin, right=thin, top=thin, bottom=thin)


def apply_grid(ws, min_row: int, max_row: int, min_col: int, max_col: int) -> None:
    thin = Side(style="thin", color=BORDER_COLOR)
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=False)


def best_width(header: str, values: pd.Series) -> float:
    text_lengths = [len(str(header))]
    for value in values.head(250):
        if pd.notna(value):
            text_lengths.append(len(str(value)))
    width = min(max(text_lengths) + 2, 40)
    if header in {"DetailLink", "Video"}:
        width = 18
    elif header in {"KeyToSymbol", "Sgp", "Measurement"}:
        width = min(max(width, 18), 32)
    elif header in {"Layout_ID", "SerialNo", "Carat", "Color", "Clarity", "Flo", "Lab"}:
        width = min(max(width, 10), 14)
    return float(max(width, 9))


def make_party_output_frame(
    rows: pd.DataFrame,
    layout_output_spec: list[tuple[str, str]],
    include_match_columns: bool,
) -> pd.DataFrame:
    # Layout columns in the merged dataframe carry the _Layout suffix when they
    # overlap with history columns. Recover them in the exact original Layout Stock
    # order and with the exact original header text.
    data: dict[str, pd.Series] = {}
    for output_header, internal_col in layout_output_spec:
        suffixed = f"{internal_col}_Layout"
        if suffixed in rows.columns:
            data[output_header] = rows[suffixed]
        elif internal_col in rows.columns:
            data[output_header] = rows[internal_col]
        else:
            data[output_header] = pd.Series([pd.NA] * len(rows), index=rows.index)

    out = pd.DataFrame(data, index=rows.index).copy()
    out.insert(0, "Sr.No", range(1, len(out) + 1))

    if include_match_columns:
        out["Match Quality"] = rows["Match Quality"].values
        out["Color Diff"] = rows["Color Diff"].values
        out["Clarity Diff"] = rows["Clarity Diff"].values
        out["Flo Diff"] = rows["Flo Diff"].values
        out["Matched History Type"] = rows["_ActionType"].values
        out["Action Weight"] = rows["_ActionWeight"].values
        out["Matched History Date"] = rows["_MatchDate"].values
        out["Matched History Serial No"] = rows["HistorySerialNo"].values
        out["Matched Stone Count"] = rows["Matched Stone Count"].values
        out["Layout Stone Count"] = rows["Layout Stone Count"].values
        out["Preference %"] = rows["Preference %"].values
        out["Minimum Required Stones"] = rows["Minimum Required Stones"].values
        out["Recommendation Score"] = rows["Recommendation Score"].values

    return out.reset_index(drop=True)


def excel_value(value: object) -> object:
    if pd.isna(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    # Convert NumPy scalar values to native Python types when available.
    item = getattr(value, "item", None)
    if callable(item):
        try:
            return item()
        except (TypeError, ValueError):
            pass
    return value


def write_party_sheet(ws, frame: pd.DataFrame) -> None:
    for col_index, header in enumerate(frame.columns, start=1):
        ws.cell(row=1, column=col_index, value=str(header))

    for row_index, (_, record) in enumerate(frame.iterrows(), start=2):
        quality = str(record.get("Match Quality", ""))
        fill_color = {
            "Exact": EXACT_FILL,
            "Nearby": NEARBY_FILL,
            "Multi Nearby": WHITE_FILL,
        }.get(quality, WHITE_FILL)
        row_fill = PatternFill("solid", fgColor=fill_color)

        for col_index, header in enumerate(frame.columns, start=1):
            value = excel_value(record[header])
            cell = ws.cell(row=row_index, column=col_index, value=value)
            cell.fill = row_fill

            if header in {"DetailLink", "Video"} and isinstance(value, str) and value.startswith("http"):
                cell.hyperlink = value
                cell.font = Font(color="0563C1", underline="single")

            if header == "Preference %" and value is not None:
                cell.number_format = "0.00%"
            elif header == "Recommendation Score" and value is not None:
                cell.number_format = "0.00"
            elif header == "Matched History Date" and value is not None:
                cell.number_format = "yyyy-mm-dd"

    max_row = max(ws.max_row, 2)
    max_col = ws.max_column
    apply_header_style(ws, 1, max_col)
    apply_grid(ws, 2, max_row, 1, max_col)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(max_col)}{max_row}"
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 30

    for col_index, header in enumerate(frame.columns, start=1):
        ws.column_dimensions[get_column_letter(col_index)].width = best_width(header, frame[header])

    # Back to navigation link at the far right.
    back_col = max_col + 2
    ws.cell(row=1, column=back_col, value="Back to Navigation")
    ws.cell(row=1, column=back_col).hyperlink = "#'Navigation'!A1"
    ws.cell(row=1, column=back_col).font = Font(color="0563C1", underline="single", bold=True)
    ws.column_dimensions[get_column_letter(back_col)].width = 20


def write_navigation_sheet(
    ws,
    salesperson: str,
    party_records: list[dict[str, object]],
    config: Config,
) -> None:
    headers = [
        "Go To Sheet", "Party Name", "Suggested Layouts", "Matched Stones",
        "Exact", "Nearby", "Other", "Country", "City", "Go →",
    ]
    for col, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col, value=header)

    if party_records:
        for row_idx, record in enumerate(party_records, start=2):
            sheet_name = str(record["Sheet Name"])
            values = [
                record["Party Name"], record["Party Name"], record["Suggested Layouts"],
                record["Matched Stones"], record["Exact"], record["Nearby"],
                record["Other"], record["Country"], record["City"], "Go →",
            ]
            for col_idx, value in enumerate(values, start=1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.fill = PatternFill("solid", fgColor=NAV_FILL)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=row_idx, column=1).hyperlink = f"#'{sheet_name}'!A1"
            ws.cell(row=row_idx, column=1).font = Font(color="0563C1", underline="single")
            ws.cell(row=row_idx, column=10).hyperlink = f"#'{sheet_name}'!A1"
            ws.cell(row=row_idx, column=10).font = Font(color="0563C1", underline="single")
    else:
        ws.merge_cells("A2:J3")
        ws["A2"] = "No suggested parties for the current data and thresholds."
        ws["A2"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws["A2"].fill = PatternFill("solid", fgColor="F2F2F2")
        ws["A2"].font = Font(italic=True, color="666666")

    apply_header_style(ws, 1, len(headers))
    apply_grid(ws, 2, max(ws.max_row, 3), 1, len(headers))
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 30

    widths = [30, 30, 18, 16, 10, 10, 10, 16, 16, 10]
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    # Legend and parameters.
    ws["L1"] = "Match Color Legend"
    ws["L1"].fill = PatternFill("solid", fgColor=HEADER_FILL)
    ws["L1"].font = Font(bold=True, color=HEADER_FONT)
    ws["L2"] = "Exact: all Color, Clarity and Flo exact"
    ws["L2"].fill = PatternFill("solid", fgColor=EXACT_FILL)
    ws["L3"] = "Nearby: exactly one variable within ±1"
    ws["L3"].fill = PatternFill("solid", fgColor=NEARBY_FILL)
    ws["L4"] = "Other: two or three variables within ±1"
    ws["L4"].fill = PatternFill("solid", fgColor=WHITE_FILL)
    ws["L6"] = "Salesperson"
    ws["M6"] = salesperson
    ws["L7"] = "Tolerance"
    ws["M7"] = config.tolerance
    ws["L8"] = "Minimum layout %"
    ws["M8"] = config.min_layout_pct
    ws["M8"].number_format = "0%"
    ws["L9"] = "Minimum stones"
    ws["M9"] = config.min_absolute_stones
    ws["L10"] = "Minimum action weight"
    ws["M10"] = config.min_action_weight
    ws.column_dimensions["L"].width = 42
    ws.column_dimensions["M"].width = 22


def create_salesperson_workbook(
    salesperson: str,
    salesperson_rows: pd.DataFrame,
    salesperson_summary: pd.DataFrame,
    layout_output_spec: list[tuple[str, str]],
    config: Config,
    output_path: Path,
) -> None:
    wb = Workbook()
    nav = wb.active
    nav.title = "Navigation"
    used_sheet_names = {"navigation"}
    party_records: list[dict[str, object]] = []

    if not salesperson_rows.empty:
        parties = sorted(str(party) for party in salesperson_rows["SoldParty"].dropna().unique())
        for party in parties:
            party_rows = salesperson_rows[salesperson_rows["SoldParty"] == party].copy()
            party_rows = party_rows.sort_values(
                by=["Layout_ID", "Segment", "_ChangedVariables", "_TotalDiff", "_LayoutSerial"],
                ascending=[True, True, True, True, True],
            )
            sheet_name = safe_sheet_name(party, used_sheet_names)
            ws = wb.create_sheet(sheet_name)
            output_frame = make_party_output_frame(
                party_rows,
                layout_output_spec,
                include_match_columns=config.include_match_columns,
            )
            write_party_sheet(ws, output_frame)

            party_summary = salesperson_summary[salesperson_summary["SoldParty"] == party]
            country = next((str(v) for v in party_summary["Country"] if pd.notna(v) and str(v).strip()), "")
            city = next((str(v) for v in party_summary["City"] if pd.notna(v) and str(v).strip()), "")
            party_records.append(
                {
                    "Sheet Name": sheet_name,
                    "Party Name": party,
                    "Suggested Layouts": int(party_summary["Layout_ID"].nunique()),
                    "Matched Stones": int(len(output_frame)),
                    "Exact": int((party_rows["Match Quality"] == "Exact").sum()),
                    "Nearby": int((party_rows["Match Quality"] == "Nearby").sum()),
                    "Other": int((party_rows["Match Quality"] == "Multi Nearby").sum()),
                    "Country": country,
                    "City": city,
                }
            )

    write_navigation_sheet(nav, salesperson, party_records, config)
    wb.save(output_path)


def create_run_summary(
    path: Path,
    config: Config,
    layout: pd.DataFrame,
    history: pd.DataFrame,
    all_summary: pd.DataFrame,
    suggested_rows: pd.DataFrame,
) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Run Summary"

    rows = [
        ["Parameter", "Value"],
        ["Tolerance", config.tolerance],
        ["Minimum Layout %", config.min_layout_pct],
        ["Minimum Absolute Stones", config.min_absolute_stones],
        ["Minimum Action Weight", config.min_action_weight],
        ["Layout Rows", len(layout)],
        ["History Rows", len(history)],
        ["Salespersons", history["Salesperson"].nunique(dropna=True)],
        ["Parties", history["SoldParty"].nunique(dropna=True)],
        ["Layouts", layout["Layout_ID"].nunique(dropna=True)],
        ["Suggested Party-Layouts", 0 if all_summary.empty else int(all_summary["Suggested"].sum())],
        ["Suggested Matched Stones", len(suggested_rows)],
    ]
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value)
    apply_header_style(ws, 1, 2)
    apply_grid(ws, 2, len(rows), 1, 2)
    ws["B3"].number_format = "0%"
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 20
    ws.sheet_view.showGridLines = False

    if not all_summary.empty:
        detail = wb.create_sheet("Party Layout Summary")
        export = all_summary.copy()
        export["Latest Match Date"] = pd.to_datetime(export["Latest Match Date"], errors="coerce")
        for c_idx, header in enumerate(export.columns, start=1):
            detail.cell(row=1, column=c_idx, value=str(header))
        for r_idx, (_, record) in enumerate(export.iterrows(), start=2):
            for c_idx, header in enumerate(export.columns, start=1):
                value = excel_value(record[header])
                cell = detail.cell(row=r_idx, column=c_idx, value=value)
                if header == "Preference %" and value is not None:
                    cell.number_format = "0.00%"
                elif header == "Recommendation Score" and value is not None:
                    cell.number_format = "0.00"
                elif header == "Latest Match Date" and value is not None:
                    cell.number_format = "yyyy-mm-dd"
        apply_header_style(detail, 1, detail.max_column)
        apply_grid(detail, 2, max(detail.max_row, 2), 1, detail.max_column)
        detail.freeze_panes = "A2"
        detail.auto_filter.ref = f"A1:{get_column_letter(detail.max_column)}{detail.max_row}"
        detail.sheet_view.showGridLines = False
        for col in range(1, detail.max_column + 1):
            detail.column_dimensions[get_column_letter(col)].width = 18

    wb.save(path)


def build_outputs(
    layout_path: Path,
    sales_path: Path,
    actions_path: Path,
    output_zip: Path,
    config: Config,
) -> dict[str, int]:
    layout_raw = read_first_nonempty_sheet(layout_path)
    sales_raw = read_first_nonempty_sheet(sales_path)
    actions_raw = read_first_nonempty_sheet(actions_path)

    layout, layout_output_spec = standardize_layout(layout_raw)
    sales = standardize_history(sales_raw, "Sales")
    actions = standardize_history(actions_raw, "Actions")
    history = pd.concat([sales, actions], ignore_index=True, sort=False)

    # Remove unusable party/salesperson records.
    history = history[
        history["SoldParty"].notna()
        & history["Salesperson"].notna()
    ].copy()

    best_matches = create_matches(layout, history, config)
    all_summary, suggested_rows = build_recommendations(layout, best_matches, config)

    output_zip.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="layout_suggestions_") as temp_dir_name:
        temp_dir = Path(temp_dir_name)
        generated_files: list[Path] = []

        salespersons = sorted(str(x) for x in history["Salesperson"].dropna().unique())
        used_filenames: set[str] = set()
        for salesperson in salespersons:
            display_name = salesperson_display_name(salesperson)
            if suggested_rows.empty:
                person_rows = pd.DataFrame()
                person_summary = pd.DataFrame()
            else:
                person_rows = suggested_rows[suggested_rows["Salesperson"] == salesperson].copy()
                person_summary = all_summary[
                    (all_summary["Salesperson"] == salesperson)
                    & (all_summary["Suggested"])
                ].copy()

            if person_rows.empty and not config.create_empty_salesperson_files:
                continue

            file_stem = safe_filename(display_name)
            candidate = file_stem
            duplicate_number = 2
            while candidate.casefold() in used_filenames:
                candidate = f"{file_stem} ({duplicate_number})"
                duplicate_number += 1
            used_filenames.add(candidate.casefold())
            file_path = temp_dir / f"{candidate}.xlsx"
            create_salesperson_workbook(
                display_name,
                person_rows,
                person_summary,
                layout_output_spec,
                config,
                file_path,
            )
            generated_files.append(file_path)

        summary_path = temp_dir / "00_Run_Summary.xlsx"
        create_run_summary(summary_path, config, layout, history, all_summary, suggested_rows)
        generated_files.append(summary_path)

        with zipfile.ZipFile(output_zip, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for file_path in generated_files:
                zf.write(file_path, arcname=file_path.name)

    return {
        "layout_rows": len(layout),
        "history_rows": len(history),
        "base_match_rows": len(best_matches),
        "suggested_rows": len(suggested_rows),
        "salesperson_files": len(salespersons),
    }


# -----------------------------------------------------------------------------
# Google Drive path handling and command line
# -----------------------------------------------------------------------------


def mount_google_drive_if_needed(project_dir: Path) -> None:
    """Mount Google Drive automatically when running inside Colab."""
    if not str(project_dir).startswith("/content/drive/"):
        return
    try:
        from google.colab import drive  # type: ignore
    except ImportError:
        return

    drive_root = Path("/content/drive/MyDrive")
    if not drive_root.exists():
        drive.mount("/content/drive")


def resolve_project_file(project_dir: Path, value: Optional[Path]) -> Optional[Path]:
    """Resolve an optional absolute path or a filename relative to project_dir."""
    if value is None:
        return None
    value = Path(value)
    return value if value.is_absolute() else project_dir / value


def candidate_xlsx_files(project_dir: Path) -> list[Path]:
    """Return input candidates while excluding generated output folders/files."""
    if not project_dir.exists():
        return []

    candidates: list[Path] = []
    for path in project_dir.glob("*.xlsx"):
        name = path.name.casefold()
        if name.startswith("~$"):
            continue
        if name == "00_run_summary.xlsx":
            continue
        candidates.append(path)
    return sorted(candidates, key=lambda p: p.name.casefold())


def auto_detect_input(project_dir: Path, role: str) -> Path:
    """
    Detect one workbook by filename.

    role values:
        layout, sales, actions
    """
    files = candidate_xlsx_files(project_dir)
    if not files:
        raise FileNotFoundError(
            f"No XLSX files were found in project folder: {project_dir}"
        )

    role_terms = {
        "layout": ("layout", "stock"),
        "sales": ("sale", "sales"),
        "actions": ("action", "actions"),
    }
    exclude_terms = {
        "layout": ("sale", "action"),
        "sales": ("action", "layout"),
        "actions": ("sale", "layout"),
    }

    terms = role_terms[role]
    excluded = exclude_terms[role]

    scored: list[tuple[int, Path]] = []
    for path in files:
        name = path.stem.casefold()
        if any(term in name for term in excluded):
            continue
        score = sum(1 for term in terms if term in name)
        if role == "layout" and "layout stock" in name:
            score += 3
        if role == "sales" and re.search(r"\bsales?\b", name):
            score += 3
        if role == "actions" and re.search(r"\bactions?\b", name):
            score += 3
        if score > 0:
            scored.append((score, path))

    if not scored:
        available = "\n  - ".join(path.name for path in files)
        raise FileNotFoundError(
            f"Could not auto-detect the {role} workbook in {project_dir}.\n"
            f"Available XLSX files:\n  - {available}\n"
            f"Pass the exact filename using --{role}."
        )

    scored.sort(key=lambda item: (-item[0], item[1].name.casefold()))
    top_score = scored[0][0]
    top = [path for score, path in scored if score == top_score]
    if len(top) > 1:
        choices = "\n  - ".join(path.name for path in top)
        raise RuntimeError(
            f"More than one possible {role} workbook was found:\n  - {choices}\n"
            f"Pass the exact filename using --{role}."
        )
    return top[0]


def resolve_input_file(
    project_dir: Path,
    supplied: Optional[Path],
    role: str,
) -> Path:
    path = resolve_project_file(project_dir, supplied)
    if path is None:
        path = auto_detect_input(project_dir, role)
    if not path.exists():
        raise FileNotFoundError(path)
    return path


def resolve_output_zip(
    project_dir: Path,
    output_folder: str,
    supplied: Optional[Path],
) -> Path:
    if supplied is not None:
        supplied = Path(supplied)
        return supplied if supplied.is_absolute() else project_dir / supplied
    return project_dir / output_folder / DEFAULT_OUTPUT_ZIP


def parse_args(argv: Optional[list[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Create salesperson-wise layout suggestion workbooks from a Google Drive project folder."
    )
    parser.add_argument(
        "--project-dir",
        type=Path,
        default=DEFAULT_PROJECT_DIR,
        help=f"Project folder. Default: {DEFAULT_PROJECT_DIR}",
    )
    parser.add_argument(
        "--layout",
        type=Path,
        default=None,
        help="Layout Stock XLSX filename inside project-dir, or an absolute path. Auto-detected when omitted.",
    )
    parser.add_argument(
        "--sales",
        type=Path,
        default=None,
        help="Sales XLSX filename inside project-dir, or an absolute path. Auto-detected when omitted.",
    )
    parser.add_argument(
        "--actions",
        type=Path,
        default=None,
        help="Actions XLSX filename inside project-dir, or an absolute path. Auto-detected when omitted.",
    )
    parser.add_argument("--tolerance", type=int, choices=[0, 1], default=1)
    parser.add_argument("--min-layout-pct", type=float, default=0.10)
    parser.add_argument("--min-absolute-stones", type=int, default=5)
    parser.add_argument("--min-action-weight", type=int, default=45)
    parser.add_argument(
        "--output-folder",
        default=DEFAULT_OUTPUT_FOLDER,
        help=f"Output folder inside project-dir. Default: {DEFAULT_OUTPUT_FOLDER}",
    )
    parser.add_argument(
        "--output-zip",
        type=Path,
        default=None,
        help=(
            "ZIP path. Relative paths are resolved inside project-dir. "
            f"Default: {DEFAULT_OUTPUT_FOLDER}/{DEFAULT_OUTPUT_ZIP}"
        ),
    )
    parser.add_argument(
        "--layout-columns-only",
        action="store_true",
        help="Do not append match audit columns after the layout-stock columns.",
    )
    parser.add_argument(
        "--skip-empty-salespersons",
        action="store_true",
        help="Do not create a workbook for a salesperson with no suggestions.",
    )
    return parser.parse_args(argv)


def main(argv: Optional[list[str]] = None) -> int:
    args = parse_args(argv)
    project_dir = Path(args.project_dir)

    mount_google_drive_if_needed(project_dir)
    project_dir.mkdir(parents=True, exist_ok=True)

    layout_path = resolve_input_file(project_dir, args.layout, "layout")
    sales_path = resolve_input_file(project_dir, args.sales, "sales")
    actions_path = resolve_input_file(project_dir, args.actions, "actions")
    output_zip = resolve_output_zip(
        project_dir,
        args.output_folder,
        args.output_zip,
    )

    config = Config(
        tolerance=args.tolerance,
        min_layout_pct=args.min_layout_pct,
        min_absolute_stones=args.min_absolute_stones,
        min_action_weight=args.min_action_weight,
        include_match_columns=not args.layout_columns_only,
        create_empty_salesperson_files=not args.skip_empty_salespersons,
    )

    print("Project folder:", project_dir)
    print("Layout file   :", layout_path.name)
    print("Sales file    :", sales_path.name)
    print("Actions file  :", actions_path.name)
    print("Tolerance     :", config.tolerance)
    print("Output ZIP    :", output_zip)

    stats = build_outputs(
        layout_path=layout_path,
        sales_path=sales_path,
        actions_path=actions_path,
        output_zip=output_zip,
        config=config,
    )

    print("\nCompleted successfully")
    print(f"Created: {output_zip}")
    for key, value in stats.items():
        print(f"{key}: {value}")

    # Show a clickable download link in Colab without forcing browser download.
    try:
        from IPython.display import FileLink, display  # type: ignore
        display(FileLink(str(output_zip)))
    except Exception:
        pass

    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise