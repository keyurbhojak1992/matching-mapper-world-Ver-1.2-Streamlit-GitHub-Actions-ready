#!/usr/bin/env python3
"""Dealmaker full trial pipeline.

This program implements the agreed workflow:

1. Completed SALE rows only are used to build real deal buckets.
2. Real deals are counted by unique stone serial numbers.
3. Party and deal-mode intervals determine WHEN a party is eligible.
4. Historical completed deals determine WHAT basket structure to prepare.
5. Action conversion is learned separately for every party and action type.
6. Recipe lines use joint combinations:
   Shape + Carat Group + Color Band + Clarity Band + Fluorescence Group + Lab
   plus preferred/normal/exceptional inclusion tolerances.
7. Current CART/HOLD/OFFER/WISHLIST rows are aggregated by unique serial.
8. Current stock is selected quota-by-quota using exact, preferred and relaxed
   match levels. The final output explains why every stone was selected.

The script is intentionally conservative. If data/history/stock coverage is weak,
it creates a review result instead of forcing a deal.
"""
from __future__ import annotations

import argparse
import importlib.util
import math
import os
import re
import sys
import time
import faulthandler
import zipfile
from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Any, Iterable

import numpy as np
import pandas as pd
from dealmaker_io_v2 import load_manual_mapping, read_dataset_table, serial_key_series
import dealmaker_quality_v6 as quality_domain
from dealmaker_quality_v6 import canonical_quality, quality_score as shared_quality_score, quality_text as shared_quality_text
from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

VERSION = "8.2.2"

QUALITY_ORDERS = quality_domain.QUALITY_ORDERS
QUALITY_MAPS = quality_domain.QUALITY_MAPS
QUALITY_REVERSE = quality_domain.QUALITY_REVERSE
QUALITY_VALUE_ALIASES = quality_domain.QUALITY_VALUE_ALIASES


def concat_nonempty_frames(frames: Iterable[pd.DataFrame]) -> pd.DataFrame:
    """Concatenate nonempty frames while avoiding pandas all-NA dtype warnings."""
    valid = [frame for frame in frames if isinstance(frame, pd.DataFrame) and not frame.empty]
    if not valid:
        return pd.DataFrame()
    if len(valid) == 1:
        return valid[0].copy()
    import warnings
    with warnings.catch_warnings():
        warnings.filterwarnings(
            "ignore",
            message="The behavior of DataFrame concatenation with empty or all-NA entries is deprecated.*",
            category=FutureWarning,
        )
        return pd.concat(valid, ignore_index=True, sort=False)

def configure_quality(control_path):
    """Load the control-workbook quality dictionary into this runtime."""
    global QUALITY_ORDERS, QUALITY_MAPS, QUALITY_REVERSE, QUALITY_VALUE_ALIASES
    cfg = quality_domain.configure_quality(control_path)
    QUALITY_ORDERS = cfg.orders
    QUALITY_MAPS = cfg.score_maps
    QUALITY_REVERSE = cfg.reverse_maps
    QUALITY_VALUE_ALIASES = cfg.alias_maps
    globals()["SHADE_ORDER"] = QUALITY_ORDERS["Shade"]
    globals()["MILKY_ORDER"] = QUALITY_ORDERS["Milky"]
    globals()["OPEN_ORDER"] = QUALITY_ORDERS["Open"]
    globals()["BLK_ORDER"] = QUALITY_ORDERS["Blk"]
    return cfg

# -----------------------------------------------------------------------------
# Configuration defaults
# -----------------------------------------------------------------------------
DEFAULT_SETTINGS: list[tuple[str, Any, str]] = [
    ("MIN_DEAL_STONES", 6, "Minimum unique completed-SALE stones in a real deal."),
    ("MAX_DEAL_DAYS", 2, "Maximum continuous calendar days in one real deal."),
    ("MIN_DEAL_BUCKETS_FOR_MAKER", 3, "Minimum real deals for party-specific automatic model."),
    ("MIN_INTERVALS_FOR_AUTO", 2, "Minimum measured intervals for automatic due output."),
    ("DUE_SOON_RATIO", 0.80, "DUE SOON begins at this fraction of normal interval."),
    ("DUE_NOW_RATIO", 1.00, "DUE NOW begins at this fraction of normal interval."),
    ("OVERDUE_RATIO", 1.30, "OVERDUE begins above this fraction of normal interval."),
    ("INCLUDE_DUE_SOON", "NO", "YES permits DUE SOON parties to generate a trial deal."),
    ("SALE_TYPE_VALUES", "SALE,SALES", "Comma-separated completed-sale values."),
    ("BID_TYPE_VALUES", "BID", "Comma-separated BID values; never treated as completed sales."),
    ("HOLD_TYPE_VALUES", "HOLD,ON HOLD", "Source action values mapped to canonical HOLD."),
    ("CART_TYPE_VALUES", "CART", "Source action values mapped to canonical CART."),
    ("OFFER_TYPE_VALUES", "OFFER", "Source action values mapped to canonical OFFER."),
    ("WISHLIST_TYPE_VALUES", "WISHLIST,WISH LIST", "Source action values mapped to canonical WISHLIST."),
    ("DETAILS_TYPE_VALUES", "DETAILS,DETAIL", "Source action values mapped to canonical DETAILS."),
    ("VIEW_TYPE_VALUES", "VIEW,VIEWS", "Source action values mapped to canonical VIEW."),
    ("AS_OF_DATE", "", "Optional DD-MM-YYYY override; blank uses latest sales/action date."),
    ("RECENT_DEALS_FOR_PATTERN", 5, "Recent deal count used in pattern scoring."),
    ("ACTION_CONVERSION_WINDOW_DAYS", 30, "Days after an action in which a completed sale counts as conversion."),
    ("CART_MAX_AGE_DAYS", 30, "Maximum age for current CART serial."),
    ("HOLD_MAX_AGE_DAYS", 14, "Maximum age for current HOLD serial."),
    ("OFFER_MAX_AGE_DAYS", 30, "Maximum age for current OFFER serial."),
    ("WISHLIST_MAX_AGE_DAYS", 60, "Maximum age for current WISHLIST serial."),
    ("DETAILS_MAX_AGE_DAYS", 14, "Maximum age for current DETAILS serial."),
    ("VIEW_MAX_AGE_DAYS", 7, "Maximum age for current VIEW serial."),
    ("MIN_ACTIONS_FOR_BEHAVIOR", 5, "Minimum actions of a type before its conversion behavior is trusted."),
    ("MIN_CONVERSION_FOR_TRIGGER", 0.10, "Minimum exact/pattern conversion used to call an action-driven behavior."),
    ("MIN_RECIPE_DEAL_PRESENCE", 0.15, "Minimum deal presence for an automatic recipe line."),
    ("MAX_RECIPE_LINES", 15, "Maximum preferred combination lines in one basket archetype."),
    ("RECIPE_CUMULATIVE_SHARE", 0.90, "Target historical stone share represented by recipe lines."),
    ("ALLOW_RELAXED_MATCH", "YES", "Permit relaxed matching when exact/preferred stock is unavailable."),
    ("SMALL_MIN_STOCK_COVERAGE", 0.70, "Minimum selected/target coverage for small deal generation."),
    ("MEDIUM_MIN_STOCK_COVERAGE", 0.75, "Minimum selected/target coverage for medium deal generation."),
    ("BULK_MIN_STOCK_COVERAGE", 0.80, "Minimum selected/target coverage for bulk deal generation."),
    ("CORE_MIN_STOCK_COVERAGE", 0.90, "Minimum fill rate for core recipe lines."),
    ("MEDIUM_MIN_CURRENT_ACTION_STONES", 5, "Suggested action-linked available stones before medium mode is preferred."),
    ("BULK_MIN_CURRENT_ACTION_STONES", 11, "Suggested action-linked available stones before bulk mode is preferred."),
    ("MAX_CONTROLLED_EXPLORATION_SHARE", 0.10, "Maximum flexible/exploration share in generated basket."),
    ("UNKNOWN_QUALITY_POLICY", "EXCLUDE", "REVIEW or EXCLUDE for unknown stock quality values. v6 defaults to fail-closed EXCLUDE."),
    ("QUALITY_DOMAIN_POLICY", "ERROR", "ERROR blocks production when an unmapped quality-domain value is found; WARNING permits review only."),
    ("PAIR_MODE", "COMPLETE_PAIR", "When either eligible pair member is selected, automatically add the related member to the same party basket."),
    ("PAIR_ALLOCATION_POLICY", "RESERVE_PAIR_KEY", "Pairs are allocated as one unit; both serials are reserved to the same accepted party."),
    ("PAIR_EXCLUSION_POLICY", "CASCADE", "CASCADE removes the complete pair when either member fails stock, quality, completed-sale, or business exclusion rules."),
    ("DIRECT_ACTION_POLICY", "DIRECT_FIRST_WITHIN_MATCH", "DIRECT_FIRST_WITHIN_MATCH prioritizes valid HOLD/CART/OFFER stones within the same match level; SCORE_WEIGHTED uses score only."),
    ("OUTPUT_ONLY_INTERVAL_ELIGIBLE", "YES", "YES generates deals only for interval-eligible parties."),
    ("REQUIRE_STOCK_STATUS_COLUMN", "NO", "YES stops production when no stock availability/status column is mapped."),
    ("STOCK_SNAPSHOT_CONFIRMED_LIVE", "NO", "Set YES only when the uploaded stock file is confirmed to contain live available inventory only. Without status or this confirmation, v6 blocks READY release status."),
    ("STOCK_AVAILABLE_VALUES", "AVAILABLE,IN STOCK,STOCK,YES,Y,1,ACTIVE,OPEN", "Values treated as available when Stock_Status exists."),
    ("STOCK_UNAVAILABLE_VALUES", "SOLD,RESERVED,UNAVAILABLE,NO,N,0,INACTIVE,MEMO", "Values removed from current stock when Stock_Status exists."),
    ("STOCK_UNKNOWN_STATUS_POLICY", "REVIEW", "REVIEW keeps unknown status rows flagged; EXCLUDE removes them; ERROR stops the run."),
    ("EXCLUDE_COMPLETED_SALE_SERIALS_FROM_STOCK", "YES", "YES removes stock serials found in completed SALE history unless the source history is intentionally reusable."),
    ("PRIMARY_TRIGGER_ACTION_TYPES", "HOLD,CART,OFFER,WISHLIST", "Only these action types may define the primary buyer behavior; DETAILS and VIEW remain supporting signals."),
    ("DIRECT_SELECTION_ACTION_TYPES", "HOLD,CART,OFFER", "Only these current actions label a selected stone as direct interest; WISHLIST, DETAILS and VIEW remain supporting signals."),
    ("GLOBAL_STOCK_ALLOCATION", "UNIQUE", "UNIQUE reserves accepted serials and pair keys globally in deterministic urgency order. INDEPENDENT is for comparison only."),
    ("MAX_CANDIDATE_ROWS_IN_WORKBOOK", 50000, "Maximum All_Stock_Candidates rows written to the analysis workbook; the complete table is exported separately when larger."),
    ("MODE_DUE_REQUIRED_FOR_SELECTION", "YES", "YES permits automatic selection only for modes whose own cadence is DUE NOW or OVERDUE (or DUE SOON when enabled)."),
    ("ALLOW_CROSS_MODE_FALLBACK", "NO", "NO prevents an early mode from being used as a fallback when a due mode is unfillable."),
    ("MIN_MODE_INTERVAL_COUNT_FOR_INDEPENDENT_CADENCE", 2, "Minimum same-mode intervals required before the mode cadence is treated as independently reliable."),
    ("ACTION_POLICY", "BOOST_ONLY", "BOOST_ONLY uses Actions only for ranking, confidence, and review labels; completed Sales history controls eligibility, timing, mode, target, and recipe."),
    ("NO_ACTION_RESULT", "GENERATE_WITH_REVIEW", "A due historically proven basket with no recent Actions may generate when all sales-history, stock, quality, pair, and allocation controls pass; label it for review."),
    ("STALE_ACTION_RESULT", "GENERATE_WITH_REVIEW", "Stale Actions do not reject a valid due basket; they create an explicit review label and remain subject to source-freshness release blocking."),
    ("ACTIONS_CAN_OVERRIDE_EARLY_MODE", "NO", "NO means Actions can never activate a mode whose independent Sales cadence is early."),
    ("ACTIONS_CAN_OVERRIDE_STOCK_SHORTAGE", "NO", "NO means Actions can never override recipe fill, stock coverage, quality, pair, or unique-allocation controls."),
    ("MEDIUM_ACTION_REQUIREMENT_IS_HARD_GATE", "NO", "Legacy compatibility control. Under ACTION_POLICY=BOOST_ONLY, MEDIUM action thresholds are support indicators, not rejection gates."),
    ("BULK_ACTION_REQUIREMENT_IS_HARD_GATE", "NO", "Legacy compatibility control. Under ACTION_POLICY=BOOST_ONLY, BULK action thresholds are support indicators, not rejection gates."),
    ("ENFORCE_ABSOLUTE_MODE_MINIMUM", "YES", "YES requires at least 6 SMALL, 12 MEDIUM, or 25 BULK stones in an accepted basket."),
    ("DEFAULT_HISTORICAL_SHADE_LIMIT", "NONE", "Conservative fallback when a historical profile has no known Shade observations."),
    ("DEFAULT_HISTORICAL_MILKY_LIMIT", "NONE", "Conservative fallback when a historical profile has no known Milky observations."),
    ("DEFAULT_HISTORICAL_OPEN_LIMIT", "NONE", "Conservative fallback when a historical profile has no known Open observations."),
    ("DEFAULT_HISTORICAL_BLK_LIMIT", "NONE", "Conservative fallback when a historical profile has no known black-inclusion observations."),
    ("MAX_SALES_STALENESS_DAYS", 3, "Maximum Sales source age before commercial release requires a fresh export or audited override."),
    ("MAX_ACTIONS_STALENESS_DAYS", 7, "Maximum Actions source age before commercial release requires a fresh export or audited override."),
    ("RELEASE_REQUIRES_SOURCE_FRESHNESS", "YES", "YES keeps generated baskets commercially blocked when Sales or Actions exceed freshness limits."),
]

ACTION_TYPES = ["HOLD", "CART", "OFFER", "WISHLIST", "DETAILS", "VIEW"]
ACTION_BASE_WEIGHTS = {
    "HOLD": 6.0,
    "CART": 6.0,
    "OFFER": 4.0,
    "WISHLIST": 2.0,
    "DETAILS": 1.0,
    "VIEW": 0.5,
}
ACTION_AGE_SETTING = {
    "HOLD": "HOLD_MAX_AGE_DAYS",
    "CART": "CART_MAX_AGE_DAYS",
    "OFFER": "OFFER_MAX_AGE_DAYS",
    "WISHLIST": "WISHLIST_MAX_AGE_DAYS",
    "DETAILS": "DETAILS_MAX_AGE_DAYS",
    "VIEW": "VIEW_MAX_AGE_DAYS",
}
MODE_ORDER = ["SMALL", "MEDIUM", "BULK"]
MODE_LABEL = {"SMALL": "SMALL 6-11", "MEDIUM": "MEDIUM 12-24", "BULK": "BULK 25+"}
MODE_MIN_STONES = {"SMALL": 6, "MEDIUM": 12, "BULK": 25}
MODE_MAX_STONES = {"SMALL": 11, "MEDIUM": 24, "BULK": None}
MODE_COVERAGE_SETTING = {
    "SMALL": "SMALL_MIN_STOCK_COVERAGE",
    "MEDIUM": "MEDIUM_MIN_STOCK_COVERAGE",
    "BULK": "BULK_MIN_STOCK_COVERAGE",
}

COLOR_ORDER = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O-P", "Q-R", "S-T", "U-V", "W-X", "Y-Z"]
CLARITY_ORDER = ["FL", "IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2", "SI3", "I1", "I2", "I3", "I4", "I5"]
FLO_ORDER = ["NON", "FNT", "MED", "STG", "VST"]
SHADE_ORDER = QUALITY_ORDERS["Shade"]
MILKY_ORDER = QUALITY_ORDERS["Milky"]
OPEN_ORDER = QUALITY_ORDERS["Open"]
BLK_ORDER = QUALITY_ORDERS["Blk"]

STOCK_ALIASES = {
    "Loc": ["Loc", "Location", "Loc."],
    "Stock_Status": ["Stock_Status", "Stock Status", "Availability", "Availability Status", "Available", "Status"],
    "Sr.No": ["Sr.No", "Sr.NO", "Sr No"],
    "Serial No": ["Serial No", "Stone ID", "SerialNo"],
    "DETAIL": ["DETAIL", "Detail"],
    "PHYGITAL": ["PHYGITAL"],
    "Shp": ["Shp", "Shape"],
    "Carat": ["Carat"],
    "Carat_grp": ["Carat_grp", "carat_grp", "Carat Group"],
    "Col": ["Col", "Color name", "Color"],
    "Qua": ["Qua", "Clarity"],
    "Cut": ["Cut", "cut"],
    "Pol": ["Pol", "polish"],
    "Sym": ["Sym"],
    "Flo": ["Flo", "Flou.", "Fluorescence"],
    "Lab": ["Lab"],
    "ReportNo": ["ReportNo", "Certificate No"],
    "RapRate": ["RapRate", "Rap", "RAP"],
    "Reserve%": ["Reserve%"],
    "Disc%": ["Disc%", "Dis%", "Discount"],
    "Rate": ["Rate", "Price/Ct$", "$/Cts.", "Price/carat", "RATE"],
    "Amount": ["Amount", "Total Amount", "Amount$", "Total", "AMT"],
    "SGPComment": ["SGPComment", "SGP Comment"],
    "Measurement": ["Measurement", "Measurment"],
    "Ratio": ["Ratio"],
    "Dep": ["Dep", "Depth%", "Depth"],
    "Tab": ["Tab", "Table%", "Table"],
    "Dia": ["Dia"],
    "Girdle": ["Girdle", "Girdle %"],
    "C/A": ["C/A", "CR Angle"],
    "C/H": ["C/H", "CR Height"],
    "P/A": ["P/A", "Pav Angle"],
    "P/H": ["P/H", "Pav Height"],
    "EYE Clean": ["EYE Clean"],
    "Luster": ["Luster"],
    "E.Facet": ["E.Facet"],
    "Culet": ["Culet"],
    "Table Incl": ["Table Incl", "Table Inclusion"],
    "SINC_NAME": ["SINC_NAME", "Side Inclusion"],
    "TOIN_NAME": ["TOIN_NAME", "Table Open"],
    "SOIN_NAME": ["SOIN_NAME", "Side Open"],
    "GIRDLE_OPEN_NAME": ["GIRDLE_OPEN_NAME", "Girdle Open"],
    "PAV_OPEN_NAME": ["PAV_OPEN_NAME", "Pavilion Open"],
    "TBINC_NAME": ["TBINC_NAME", "Table Black", "Table Black Incl."],
    "SBINC_NAME": ["SBINC_NAME", "Side Black", "Side Black Incl."],
    "Spot Incl.": ["Spot Incl."],
    "KeyToSymbol": ["KeyToSymbol", "Key To Symbol"],
    "Comment": ["Comment", "comment"],
    "Remarks": ["Remarks"],
    "IMAGE": ["IMAGE"],
    "VIDEO": ["VIDEO"],
    "CERTIFICATE": ["CERTIFICATE"],
    "Shade": ["Shade"],
    "Milky": ["Milky"],
}

OUTPUT_STOCK_COLUMNS = [
    "Selection_Rank", "Deal_Mode", "Recipe_Line_ID", "Recipe_Priority", "Selection_Source",
    "Primary_Action_Type", "Last_Action_Date", "Action_Age_Days", "Action_Count",
    "Party_Action_Conversion", "Direct_Action_Flag", "Match_Level", "Historical_Combination_Score",
    "Quality_Fit_Score", "Discount_Fit_Score", "Final_Selection_Score", "Selection_Reason",
    "Pair_Key", "Pair_Review_Flag", "Quality_Review_Flag", "Sr.No", "Serial No", "DETAIL",
    "PHYGITAL", "Shp", "Carat", "Carat_grp", "Col", "Qua", "Cut", "Pol", "Sym", "Flo",
    "Lab", "ReportNo", "RapRate", "Reserve%", "Disc%", "Rate", "Amount", "SGPComment", "Loc",
    "Measurement", "Ratio", "Dep", "Tab", "Dia", "Girdle", "C/A", "C/H", "P/A", "P/H",
    "EYE Clean", "Luster", "E.Facet", "Culet", "Table Incl", "SINC_NAME", "TOIN_NAME",
    "SOIN_NAME", "GIRDLE_OPEN_NAME", "PAV_OPEN_NAME", "TBINC_NAME", "SBINC_NAME", "Spot Incl.",
    "KeyToSymbol", "Comment", "Remarks", "IMAGE", "VIDEO", "CERTIFICATE", "Shade", "Milky",
]


# -----------------------------------------------------------------------------
# Generic helpers
# -----------------------------------------------------------------------------
def clean_text(value: Any, default: str = "") -> str:
    if pd.isna(value):
        return default
    return str(value).strip()


def yes(value: Any) -> bool:
    return clean_text(value).upper() in {"YES", "Y", "TRUE", "1"}



def first_nonblank(series: pd.Series) -> str:
    """Return the first non-empty value from a Series as clean text."""
    if series is None:
        return ""
    for value in series:
        text = clean_text(value)
        if text:
            return text
    return ""

def parse_csv_set(value: Any) -> set[str]:
    return {x.strip().upper() for x in str(value).split(",") if x.strip()}


def canonicalize_action_types(actions: pd.DataFrame, settings: dict[str, Any]) -> pd.DataFrame:
    out = actions.copy()
    mapping: dict[str, str] = {}
    for canonical in ACTION_TYPES:
        for source in parse_csv_set(settings.get(f"{canonical}_TYPE_VALUES", canonical)):
            mapping[source] = canonical
    out["Original_Action_Type"] = out["Type"]
    out["Type"] = out["Type"].astype(str).str.strip().str.upper().map(lambda x: mapping.get(x, x))
    return out


def normalize_upper(series: pd.Series, default: str = "UNKNOWN") -> pd.Series:
    return series.fillna(default).astype(str).str.strip().str.upper().replace("", default)


def safe_excel(value: Any) -> Any:
    if pd.isna(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    if isinstance(value, np.generic):
        return value.item()
    return value


def load_module(path: Path, module_name: str):
    spec = importlib.util.spec_from_file_location(module_name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Could not load module: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


def mode_for_stones(value: Any) -> str:
    try:
        n = int(round(float(value)))
    except Exception:
        return "UNKNOWN"
    if n <= 11:
        return "SMALL"
    if n <= 24:
        return "MEDIUM"
    return "BULK"


def color_band(value: Any) -> str:
    v = clean_text(value, "UNKNOWN").upper()
    if v in {"D", "E", "F"}:
        return "D-F"
    if v in {"G", "H"}:
        return "G-H"
    if v in {"I", "J"}:
        return "I-J"
    if v in {"K", "L", "M", "N"}:
        return "K-N"
    if v in {"O-P", "Q-R", "S-T", "U-V", "W-X", "Y-Z"}:
        return "O-Z"
    return v


def clarity_band(value: Any) -> str:
    v = clean_text(value, "UNKNOWN").upper()
    if v in {"FL", "IF"}:
        return "FL-IF"
    if v.startswith("VVS"):
        return "VVS"
    if v.startswith("VS"):
        return "VS"
    if v.startswith("SI"):
        return "SI"
    if v.startswith("I"):
        return "I"
    return v


def flo_group(value: Any) -> str:
    v = clean_text(value, "UNKNOWN").upper()
    if v in {"VSL", "SL"}:
        v = "FNT"
    if v in {"NON", "FNT"}:
        return "NON-FNT"
    if v == "MED":
        return "MED"
    if v in {"STG", "VST"}:
        return "STG-VST"
    return v


def quality_score(value: Any, dimension: str) -> float:
    return shared_quality_score(value, dimension, blank_as_none=False)


def quality_text(value: Any, dimension: str) -> str:
    return shared_quality_text(value, dimension)


def row_max_quality(row: pd.Series, columns: list[str], dimension: str) -> float:
    vals = [quality_score(row.get(c), dimension) for c in columns]
    vals = [x for x in vals if pd.notna(x)]
    return max(vals) if vals else np.nan


def mode_or_median(series: pd.Series) -> float:
    s = pd.to_numeric(series, errors="coerce").dropna()
    if s.empty:
        return np.nan
    m = s.mode()
    return float(m.min()) if not m.empty else float(s.median())


def normal_max_score(series: pd.Series) -> float:
    s = pd.to_numeric(series, errors="coerce").dropna()
    if s.empty:
        return np.nan
    # 90th percentile is less sensitive to one exceptional historical stone.
    return float(min(math.ceil(float(s.quantile(0.90))), int(s.max())))


def largest_remainder(weights: pd.Series, target: int) -> pd.Series:
    if target <= 0 or len(weights) == 0:
        return pd.Series(0, index=weights.index, dtype=int)
    w = pd.to_numeric(weights, errors="coerce").fillna(0).clip(lower=0)
    if w.sum() <= 0:
        w = pd.Series(1.0, index=w.index)
    raw = w / w.sum() * target
    base = np.floor(raw).astype(int)
    remaining = target - int(base.sum())
    if remaining > 0:
        for idx in (raw - base).sort_values(ascending=False).index[:remaining]:
            base.loc[idx] += 1
    return base.astype(int)


def clean_pair_serial(value: Any) -> str:
    """Normalize pair/serial identifiers without losing meaningful punctuation."""
    if pd.isna(value) or value == "":
        return ""
    text = str(value).strip().upper()
    text = re.sub(r"\.0+$", "", text)
    return re.sub(r"\s+", "", text)


def extract_pair_id(value: Any) -> str:
    text = clean_text(value)
    if not text:
        return ""
    match = re.search(r"(?:PAIRED\s+WITH|MATCH\s+PAIR)\s*[^\d]*([\d.]+)", text, re.I)
    return clean_pair_serial(match.group(1)) if match else ""


def get_pair_id_for_sorting(text: Any) -> str | None:
    pair_id = extract_pair_id(text)
    return pair_id or None


def annotate_pair_groups(df: pd.DataFrame) -> pd.DataFrame:
    """Build a bidirectional pair registry from one-sided SGP comments.

    The referenced serial is included in the expected group even when it is not
    present in the source. This allows the eligibility gate to fail closed for a
    missing partner and gives both observed members the same Pair_Key.
    """
    if df is None or df.empty:
        return df.copy() if isinstance(df, pd.DataFrame) else df
    serial_col = "Serial_Key" if "Serial_Key" in df.columns else ("Serial No" if "Serial No" in df.columns else None)
    if serial_col is None:
        return df.copy()

    out = df.copy()
    out["_Pair_Serial"] = out[serial_col].apply(clean_pair_serial)
    comments = out.get("SGPComment", pd.Series("", index=out.index))
    out["_Pair_Reference"] = comments.apply(extract_pair_id)

    graph: dict[str, set[str]] = {}
    for serial, partner in out[["_Pair_Serial", "_Pair_Reference"]].itertuples(index=False, name=None):
        if not serial or not partner or serial == partner:
            continue
        graph.setdefault(serial, set()).add(partner)
        graph.setdefault(partner, set()).add(serial)

    serial_to_members: dict[str, tuple[str, ...]] = {}
    visited: set[str] = set()
    for serial in sorted(graph):
        if serial in visited:
            continue
        stack = [serial]
        component: set[str] = set()
        while stack:
            current = stack.pop()
            if current in visited:
                continue
            visited.add(current)
            component.add(current)
            stack.extend(graph.get(current, set()) - visited)
        members = tuple(sorted(component))
        for member in members:
            serial_to_members[member] = members

    observed = set(out["_Pair_Serial"].dropna().astype(str))
    observed.discard("")

    def members_for(serial: str) -> tuple[str, ...]:
        return serial_to_members.get(serial, tuple())

    out["Pair_Group_Members"] = out["_Pair_Serial"].map(
        lambda serial: "|".join(members_for(serial)) if members_for(serial) else ""
    )
    out["Pair_Key"] = out["_Pair_Serial"].map(
        lambda serial: "_".join(members_for(serial)) if len(members_for(serial)) >= 2 else ""
    )
    out["Pair_Partner_Serial"] = out["_Pair_Serial"].map(
        lambda serial: next((member for member in members_for(serial) if member != serial), "")
    )
    out["Pair_Source_Complete"] = out.apply(
        lambda row: "YES" if not row["Pair_Key"] or set(str(row["Pair_Group_Members"]).split("|")).issubset(observed) else "NO",
        axis=1,
    )
    out["Pair_Protection_Review_Flag"] = out.get("Pair_Protection_Review_Flag", "")
    return out.drop(columns=["_Pair_Serial", "_Pair_Reference"], errors="ignore")


def cascade_incomplete_pair_groups(
    eligible_df: pd.DataFrame,
    reason: str,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Remove every observed member of a pair when the eligible group is incomplete."""
    if eligible_df is None or eligible_df.empty or "Pair_Key" not in eligible_df.columns:
        empty = eligible_df.iloc[0:0].copy() if isinstance(eligible_df, pd.DataFrame) else pd.DataFrame()
        return eligible_df, empty

    work = eligible_df.copy()
    serials = set(work.get("Serial_Key", work.get("Serial No", pd.Series(dtype=str))).apply(clean_pair_serial))
    serials.discard("")
    incomplete_keys: set[str] = set()
    pair_rows = work[work["Pair_Key"].fillna("").astype(str).ne("")]
    for pair_key, group in pair_rows.groupby("Pair_Key", sort=False):
        members_text = next((str(x) for x in group.get("Pair_Group_Members", pd.Series(dtype=str)) if str(x).strip()), "")
        expected = {clean_pair_serial(x) for x in members_text.split("|") if clean_pair_serial(x)}
        if not expected:
            expected = {clean_pair_serial(x) for x in str(pair_key).split("_") if clean_pair_serial(x)}
        source_complete = group.get("Pair_Source_Complete", pd.Series("YES", index=group.index)).astype(str).str.upper().eq("YES").all()
        if not source_complete or not expected.issubset(serials):
            incomplete_keys.add(str(pair_key))

    remove_mask = work["Pair_Key"].fillna("").astype(str).isin(incomplete_keys)
    removed = work.loc[remove_mask].copy()
    if not removed.empty:
        removed["Pair_Protection_Review_Flag"] = reason
    return work.loc[~remove_mask].copy(), removed


def cascade_pair_exclusions(
    eligible_df: pd.DataFrame,
    directly_excluded_df: pd.DataFrame,
    reason: str,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Remove eligible pair partners when a direct exclusion removed one member."""
    if eligible_df is None or eligible_df.empty or directly_excluded_df is None or directly_excluded_df.empty:
        empty = eligible_df.iloc[0:0].copy() if isinstance(eligible_df, pd.DataFrame) else pd.DataFrame()
        return eligible_df, empty
    if "Pair_Key" not in eligible_df.columns or "Pair_Key" not in directly_excluded_df.columns:
        return eligible_df, eligible_df.iloc[0:0].copy()
    keys = set(directly_excluded_df["Pair_Key"].fillna("").astype(str))
    keys.discard("")
    mask = eligible_df["Pair_Key"].fillna("").astype(str).isin(keys)
    removed = eligible_df.loc[mask].copy()
    if not removed.empty:
        removed["Pair_Protection_Review_Flag"] = reason
    return eligible_df.loc[~mask].copy(), removed


def make_pair_key(serial: str, partner: str) -> str:
    if not serial or not partner:
        return ""
    return "_".join(sorted([clean_pair_serial(serial), clean_pair_serial(partner)]))


# -----------------------------------------------------------------------------
# Control workbook
# -----------------------------------------------------------------------------
def create_control_workbook(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Settings"
    ws.append(["Setting", "Value", "Explanation"])
    for row in DEFAULT_SETTINGS:
        ws.append(list(row))
    style_sheet(ws)
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 90

    ws = wb.create_sheet("Party_Overrides")
    ws.append(["Sold Party", "Override", "Manual_Interval_Days", "Force_Deal_Mode", "Notes"])
    ws.append(["", "AUTO", "", "AUTO", ""])
    dv1 = DataValidation(type="list", formula1='"AUTO,FORCE_INCLUDE,FORCE_EXCLUDE"', allow_blank=False)
    dv2 = DataValidation(type="list", formula1='"AUTO,SMALL,MEDIUM,BULK"', allow_blank=False)
    ws.add_data_validation(dv1); dv1.add("B2:B5000")
    ws.add_data_validation(dv2); dv2.add("D2:D5000")
    style_sheet(ws)

    ws = wb.create_sheet("Party_Name_Aliases")
    ws.append(["Source_Party", "Canonical_Party", "Active", "Notes"])
    ws.append(["", "", "YES", "Exact source name is replaced before deal and behavior analysis."])
    dv_alias = DataValidation(type="list", formula1='"YES,NO"', allow_blank=False)
    ws.add_data_validation(dv_alias); dv_alias.add("C2:C5000")
    style_sheet(ws)
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["D"].width = 70

    ws = wb.create_sheet("Removed_Parties")
    ws.append(["Sold Party", "Reason"])
    ws.append(["", ""])
    style_sheet(ws)

    ws = wb.create_sheet("Exclusion_Rules")
    ws.append([
        "Sold Party", "Shp", "Carat", "Color Name", "Qua", "Cut", "Pol", "Sym", "Flo", "Lab",
        "Disc%", "Shade", "Milky", "TOIN_NAME", "SOIN_NAME", "GIRDLE_OPEN_NAME", "PAV_OPEN_NAME",
        "TBINC_NAME", "SBINC_NAME", "Remarks",
    ])
    ws.append([""] * 20)
    style_sheet(ws)

    ws = wb.create_sheet("Column_Mapping")
    ws.append(["Dataset", "Standard_Column", "Source_Column", "Active", "Common_Aliases", "Notes"])
    mapping_rows = [
        ("SALES", "Serial No", "", "YES", "Serial No; Stone ID; SerialNo", "Required"),
        ("SALES", "Sold Party", "", "YES", "Sold Party; Party; Customer; Party Name", "Required"),
        ("SALES", "Date", "", "YES", "Date; Business Date; Sale Date", "Required"),
        ("SALES", "Type", "", "YES", "Type; Transaction Type", "Required: SALE/SALES/BID values are controlled in Settings"),
        ("SALES", "Sales_Person", "", "YES", "SALE_PERSON; Sales Person; Salesperson", "Recommended"),
        ("ACTIONS", "Serial No", "", "YES", "Serial No; Stone ID; SerialNo", "Required"),
        ("ACTIONS", "Sold Party", "", "YES", "Sold Party; Party; Customer; Party Name", "Required"),
        ("ACTIONS", "Date", "", "YES", "Date; Action Date; Business Date", "Required"),
        ("ACTIONS", "Type", "", "YES", "Type; Action Type", "Required: action aliases are controlled in Settings"),
        ("ACTIONS", "Sales_Person", "", "YES", "SALE_PERSON; Sales Person; Salesperson", "Recommended"),
        ("STOCK", "Serial No", "", "YES", "Serial No; Stone ID; SerialNo", "Required"),
        ("STOCK", "Shp", "", "YES", "Shp; Shape", "Required"),
        ("STOCK", "Carat", "", "YES", "Carat", "Required"),
        ("STOCK", "Col", "", "YES", "Col; Color; Color name", "Required"),
        ("STOCK", "Qua", "", "YES", "Qua; Clarity", "Required"),
        ("STOCK", "Flo", "", "YES", "Flo; Flou.; Fluorescence", "Required"),
        ("STOCK", "Lab", "", "YES", "Lab", "Required"),
        ("STOCK", "Stock_Status", "", "YES", "Stock Status; Availability; Available; Status", "Strongly recommended for production"),
        ("STOCK", "SGPComment", "", "YES", "SGPComment; SGP Comment", "Used for pair review"),
    ]
    for row in mapping_rows:
        ws.append(row)
    style_sheet(ws)
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 28
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 55
    ws.column_dimensions["F"].width = 55

    ws = wb.create_sheet("User_Guide")
    guide = [
        ("1", "Upload full Sales/BID, full Actions, and actual current-stock files. The production notebook does not run sample data."),
        ("2", "Use completed SALE values only for business and deal intervals."),
        ("3", "Review Column_Mapping after the preflight report. Enter Source_Column only when automatic aliases do not match."),
        ("4", "Run the preflight first, correct ERROR items, then run the production pipeline."),
        ("5", "Review Party_Generation_Status before sending any generated deal."),
        ("6", "HOLD/OFFER behavior can be learned only when those action types exist in the full action dataset."),
        ("7", "A trial stock created from an action sample proves program flow only; it is not a current availability file."),
    ]
    ws.append(["Step", "Instruction"])
    for row in guide:
        ws.append(row)
    style_sheet(ws)
    ws.column_dimensions["B"].width = 110

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def read_user_sheet(path: Path, sheet: str, required: set[str]) -> pd.DataFrame:
    try:
        first = pd.read_excel(path, sheet_name=sheet)
    except Exception:
        return pd.DataFrame(columns=sorted(required))
    if required.issubset(set(first.columns)):
        return first
    try:
        second = pd.read_excel(path, sheet_name=sheet, header=2)
        if required.issubset(set(second.columns)):
            return second
    except Exception:
        pass
    return first


def load_control(path: Path) -> tuple[dict[str, Any], pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    defaults = {k: v for k, v, _ in DEFAULT_SETTINGS}
    if not path.exists():
        create_control_workbook(path)
    settings_df = read_user_sheet(path, "Settings", {"Setting", "Value"})
    settings = defaults.copy()
    if {"Setting", "Value"}.issubset(settings_df.columns):
        for _, r in settings_df.iterrows():
            key = clean_text(r.get("Setting"))
            if key:
                settings[key] = r.get("Value")
    overrides = read_user_sheet(path, "Party_Overrides", {"Sold Party"})
    removed = read_user_sheet(path, "Removed_Parties", {"Sold Party"})
    exclusions = read_user_sheet(path, "Exclusion_Rules", {"Sold Party"})
    return settings, overrides, removed, exclusions


def load_party_aliases(path: Path) -> dict[str, str]:
    aliases = read_user_sheet(path, "Party_Name_Aliases", {"Source_Party", "Canonical_Party"})
    if aliases.empty or not {"Source_Party", "Canonical_Party"}.issubset(aliases.columns):
        return {}
    mapping: dict[str, str] = {}
    for _, row in aliases.iterrows():
        active = clean_text(row.get("Active", "YES")).upper()
        source = clean_text(row.get("Source_Party"))
        canonical = clean_text(row.get("Canonical_Party"))
        if active not in {"YES", "Y", "TRUE", "1"} or not source or not canonical:
            continue
        mapping[source.upper()] = canonical
    return mapping


def apply_party_aliases(
    df: pd.DataFrame,
    alias_map: dict[str, str],
    dataset_name: str,
    quality_log: list[dict[str, Any]],
) -> pd.DataFrame:
    if df.empty or not alias_map or "Sold Party" not in df.columns:
        return df
    out = df.copy()
    source = out["Sold Party"].fillna("").astype(str).str.strip()
    mapped = source.str.upper().map(alias_map)
    changed = mapped.notna() & mapped.ne(source)
    out.loc[mapped.notna(), "Sold Party"] = mapped[mapped.notna()]
    quality_log.append({
        "Dataset": dataset_name,
        "Check": "Party alias rows standardized",
        "Value": int(changed.sum()),
        "Severity": "INFO" if changed.any() else "OK",
    })
    return out


# -----------------------------------------------------------------------------
# Historical analysis
# -----------------------------------------------------------------------------
def add_deal_membership(sales_only: pd.DataFrame, deals: pd.DataFrame) -> pd.DataFrame:
    """Attach non-overlapping deal buckets with one party/date join."""
    if sales_only.empty or deals.empty:
        return pd.DataFrame()
    membership_rows: list[dict[str, Any]] = []
    for d in deals[["Sold Party", "Deal_ID", "Deal_Start_Date", "Deal_End_Date", "Deal_Total_Stones"]].to_dict("records"):
        start_date = d["Deal_Start_Date"]
        end_date = d["Deal_End_Date"]
        total_stones = d["Deal_Total_Stones"]
        for business_date in pd.date_range(start_date, end_date, freq="D"):
            membership_rows.append({
                "Sold Party": d["Sold Party"],
                "Date": business_date.normalize(),
                "Deal_ID": d["Deal_ID"],
                "Deal_Start_Date": start_date,
                "Deal_End_Date": end_date,
                "Deal_Total_Stones": total_stones,
                "Deal_Mode": mode_for_stones(total_stones),
            })
    membership = pd.DataFrame(membership_rows)
    out = sales_only.merge(membership, on=["Sold Party", "Date"], how="left", sort=False)
    return out[out["Deal_ID"].notna()].copy().reset_index(drop=True)

def enrich_pattern_fields(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in ["Shp", "Carat_grp", "Col", "Qua", "Flo", "Lab"]:
        if col not in out.columns:
            out[col] = "UNKNOWN"
        out[col] = normalize_upper(out[col])
    out["Color_Band"] = out["Col"].apply(color_band)
    out["Clarity_Band"] = out["Qua"].apply(clarity_band)
    out["Flo_Group"] = out["Flo"].apply(flo_group)
    out["Exact_Combination_Key"] = (
        out["Shp"] + "|" + out["Carat_grp"] + "|" + out["Col"] + "|" + out["Qua"] + "|" + out["Flo"] + "|" + out["Lab"]
    )
    out["Preferred_Combination_Key"] = (
        out["Shp"] + "|" + out["Carat_grp"] + "|" + out["Color_Band"] + "|" + out["Clarity_Band"] + "|" + out["Flo_Group"] + "|" + out["Lab"]
    )
    out["Relaxed_Combination_Key"] = (
        out["Shp"] + "|" + out["Carat_grp"] + "|" + out["Clarity_Band"] + "|" + out["Lab"]
    )
    return out


def build_mode_summary(deals: pd.DataFrame, party_summary: pd.DataFrame, settings: dict[str, Any], as_of: pd.Timestamp) -> pd.DataFrame:
    if deals.empty:
        return pd.DataFrame()
    effective_map = (
        party_summary.set_index("Sold Party")["Effective_Interval_Days"].to_dict()
        if not party_summary.empty and "Effective_Interval_Days" in party_summary.columns else {}
    )
    groups: dict[tuple[str, str], list[tuple[Any, Any, int, str]]] = {}
    party_ids: dict[str, set[str]] = {}
    cols = ["Sold Party", "Deal_Start_Date", "Deal_End_Date", "Deal_Total_Stones", "Deal_ID"]
    for party, start_date, end_date, total_stones, deal_id in deals[cols].itertuples(index=False, name=None):
        mode = mode_for_stones(total_stones)
        groups.setdefault((party, mode), []).append((start_date, end_date, int(total_stones), deal_id))
        party_ids.setdefault(party, set()).add(deal_id)

    due_soon = float(settings.get("DUE_SOON_RATIO", 0.8))
    due_now = float(settings.get("DUE_NOW_RATIO", 1.0))
    overdue = float(settings.get("OVERDUE_RATIO", 1.3))
    rows = []
    for (party, mode), values in groups.items():
        values.sort(key=lambda x: x[0])
        starts = [x[0] for x in values]
        ends = [x[1] for x in values]
        quantities = np.asarray([x[2] for x in values], dtype=float)
        intervals = [(starts[i] - starts[i - 1]).days for i in range(1, len(starts))]
        mode_interval = float(np.median(intervals)) if intervals else np.nan
        party_effective = effective_map.get(party, np.nan)
        min_mode_intervals = int(float(settings.get("MIN_MODE_INTERVAL_COUNT_FOR_INDEPENDENT_CADENCE", 2)))
        independent_mode_cadence = len(intervals) >= min_mode_intervals and pd.notna(mode_interval) and mode_interval > 0
        effective = mode_interval if independent_mode_cadence else party_effective
        interval_source = "MODE_HISTORY" if independent_mode_cadence else ("PARTY_FALLBACK" if pd.notna(party_effective) and party_effective > 0 else "NO_INTERVAL")
        last = max(ends)
        days = int((as_of - last).days) if pd.notna(last) else np.nan
        ratio = days / effective if pd.notna(days) and pd.notna(effective) and effective > 0 else np.nan
        if pd.isna(ratio):
            status = "NO INTERVAL"
        elif ratio < due_soon:
            status = "EARLY"
        elif ratio < due_now:
            status = "DUE SOON"
        elif ratio <= overdue:
            status = "DUE NOW"
        else:
            status = "OVERDUE"
        rows.append({
            "Sold Party": party,
            "Deal_Mode": mode,
            "Deal_Mode_Label": MODE_LABEL.get(mode, mode),
            "Mode_Deal_Count": len(values),
            "Mode_Interval_Count": len(intervals),
            "Mode_Median_Interval_Days": mode_interval,
            "Effective_Mode_Interval_Days": effective,
            "Mode_Interval_Source": interval_source,
            "Mode_Independent_Cadence_Reliable": "YES" if independent_mode_cadence else "NO",
            "Mode_Last_Deal_Date": last,
            "Mode_Days_Since_Last_Deal": days,
            "Mode_Interval_Progress": ratio,
            "Mode_Interval_Status": status,
            "Target_Stone_Count": int(round(float(np.median(quantities)))),
            "P25_Stone_Count": float(np.quantile(quantities, 0.25)),
            "Median_Stone_Count": float(np.median(quantities)),
            "P75_Stone_Count": float(np.quantile(quantities, 0.75)),
            "Min_Stone_Count": int(np.min(quantities)),
            "Max_Stone_Count": int(np.max(quantities)),
            "Share_of_Party_Deals": len(values) / max(len(party_ids.get(party, set())), 1),
        })
    return pd.DataFrame(rows).sort_values(["Sold Party", "Deal_Mode"]).reset_index(drop=True)

def build_action_conversion(actions: pd.DataFrame, sales_only: pd.DataFrame, deals: pd.DataFrame, window_days: int) -> pd.DataFrame:
    """Calculate party/action conversion with indexed NumPy date arrays.

    v6 builds maps only for keys that actually occur in Actions, converts dates
    to int64 once, and performs three binary searches per action.  This avoids
    large object-key as-of joins and scales reliably on the full history.
    """
    if actions.empty:
        return pd.DataFrame(columns=["Sold Party", "Action_Type"])
    acts = enrich_pattern_fields(actions).copy()
    acts = acts[acts["Type"].isin(ACTION_TYPES) & acts["Date"].notna()].copy().reset_index(drop=True)
    if acts.empty:
        return pd.DataFrame(columns=["Sold Party", "Action_Type"])
    sales = enrich_pattern_fields(sales_only).copy()
    window_ns = int(pd.Timedelta(days=window_days).value)

    def relevant_date_map(df: pd.DataFrame, key_cols: list[str], date_col: str, relevant: pd.DataFrame) -> dict[tuple, np.ndarray]:
        if df.empty or relevant.empty:
            return {}
        relevant_keys = relevant[key_cols].drop_duplicates()
        work = df[key_cols + [date_col]].merge(relevant_keys, on=key_cols, how="inner", sort=False)
        work = work.dropna(subset=[date_col]).copy()
        if work.empty:
            return {}
        work["_Date_NS"] = work[date_col].astype("int64")
        result: dict[tuple, np.ndarray] = {}
        grouper = key_cols[0] if len(key_cols) == 1 else key_cols
        for key, group in work.groupby(grouper, sort=False, dropna=False):
            if len(key_cols) == 1:
                key = (key,)
            elif not isinstance(key, tuple):
                key = (key,)
            arr = group["_Date_NS"].to_numpy(dtype=np.int64, copy=True)
            arr.sort()
            result[key] = arr
        return result

    exact_source = sales[sales["Serial_Key"].ne("")]
    print("action conversion: building exact-serial date map", flush=True)
    exact_dates = relevant_date_map(exact_source, ["Sold Party", "Serial_Key"], "Date", acts[acts["Serial_Key"].ne("")])
    print(f"action conversion: exact map keys={len(exact_dates)}", flush=True)
    print("action conversion: building preferred-pattern date map", flush=True)
    pattern_dates = relevant_date_map(sales, ["Sold Party", "Preferred_Combination_Key"], "Date", acts)
    print(f"action conversion: pattern map keys={len(pattern_dates)}", flush=True)
    print("action conversion: building party deal-date map", flush=True)
    deal_dates = relevant_date_map(deals, ["Sold Party"], "Deal_Start_Date", acts) if not deals.empty else {}
    print(f"action conversion: deal map keys={len(deal_dates)}", flush=True)

    def future_date(arr: np.ndarray | None, start_ns: int) -> int | None:
        if arr is None or arr.size == 0:
            return None
        idx = int(np.searchsorted(arr, start_ns, side="left"))
        if idx >= arr.size:
            return None
        candidate = int(arr[idx])
        return candidate if candidate <= start_ns + window_ns else None

    stats: dict[tuple[str, str], dict[str, Any]] = {}
    acts["_Date_NS"] = acts["Date"].astype("int64")
    cols = ["Sold Party", "Type", "_Date_NS", "Serial_Key", "Preferred_Combination_Key"]
    for party, action_type, start_ns, serial_raw, pref_key in acts[cols].itertuples(index=False, name=None):
        key = (party, action_type)
        rec = stats.setdefault(key, {
            "Action_Count": 0, "Unique_Serials": set(), "Exact_Count": 0,
            "Preferred_Count": 0, "Trigger_Count": 0, "Lead_Days": [],
        })
        rec["Action_Count"] += 1
        serial = clean_text(serial_raw)
        if serial:
            rec["Unique_Serials"].add(serial)
            if future_date(exact_dates.get((party, serial)), int(start_ns)) is not None:
                rec["Exact_Count"] += 1
        if future_date(pattern_dates.get((party, pref_key)), int(start_ns)) is not None:
            rec["Preferred_Count"] += 1
        next_deal = future_date(deal_dates.get((party,)), int(start_ns))
        if next_deal is not None:
            rec["Trigger_Count"] += 1
            rec["Lead_Days"].append(int((next_deal - int(start_ns)) // pd.Timedelta(days=1).value))

    rows = []
    for (party, action_type), rec in stats.items():
        action_count = int(rec["Action_Count"])
        exact_count = int(rec["Exact_Count"]); preferred_count = int(rec["Preferred_Count"]); trigger_count = int(rec["Trigger_Count"])
        exact_rate = exact_count / action_count if action_count else 0.0
        preferred_rate = preferred_count / action_count if action_count else 0.0
        trigger_rate = trigger_count / action_count if action_count else 0.0
        rows.append({
            "Sold Party": party, "Action_Type": action_type, "Action_Count": action_count,
            "Unique_Action_Serials": len(rec["Unique_Serials"]),
            "Exact_Serial_Conversions": exact_count, "Exact_Serial_Conversion_Rate": exact_rate,
            "Preferred_Pattern_Conversions": preferred_count, "Preferred_Pattern_Conversion_Rate": preferred_rate,
            "Deal_Trigger_Count": trigger_count, "Deal_Trigger_Rate": trigger_rate,
            "Median_Action_To_Deal_Days": float(np.median(rec["Lead_Days"])) if rec["Lead_Days"] else np.nan,
            "Party_Action_Conversion_Score": 0.50 * exact_rate + 0.30 * preferred_rate + 0.20 * trigger_rate,
        })
    out = pd.DataFrame(rows)
    return out.sort_values(["Sold Party", "Party_Action_Conversion_Score", "Action_Count"], ascending=[True, False, False]).reset_index(drop=True) if not out.empty else out

def build_behavior_profiles(
    party_summary: pd.DataFrame,
    mode_summary: pd.DataFrame,
    action_conversion: pd.DataFrame,
    settings: dict[str, Any],
) -> pd.DataFrame:
    rows = []
    min_actions = int(float(settings.get("MIN_ACTIONS_FOR_BEHAVIOR", 5)))
    min_conversion = float(settings.get("MIN_CONVERSION_FOR_TRIGGER", 0.10))

    ac_by_party: dict[str, list[dict[str, Any]]] = {}
    if not action_conversion.empty:
        for rec in action_conversion.to_dict("records"):
            ac_by_party.setdefault(rec.get("Sold Party", ""), []).append(rec)
    bulk_count_map: dict[str, int] = {}
    if not mode_summary.empty:
        for party, mode, count in mode_summary[["Sold Party", "Deal_Mode", "Mode_Deal_Count"]].itertuples(index=False, name=None):
            if mode == "BULK":
                bulk_count_map[party] = bulk_count_map.get(party, 0) + int(count or 0)

    for p in party_summary.to_dict("records"):
        party = p.get("Sold Party")
        ac_rows = ac_by_party.get(party, [])
        trusted = [
            r for r in ac_rows
            if int(r.get("Action_Count", 0) or 0) >= min_actions
            and float(r.get("Party_Action_Conversion_Score", 0.0) or 0.0) >= min_conversion
        ]
        trusted.sort(
            key=lambda r: (float(r.get("Party_Action_Conversion_Score", 0.0) or 0.0), int(r.get("Action_Count", 0) or 0)),
            reverse=True,
        )
        primary_types = parse_csv_set(settings.get("PRIMARY_TRIGGER_ACTION_TYPES", "HOLD,CART,OFFER,WISHLIST"))
        trusted_primary = [r for r in trusted if clean_text(r.get("Action_Type")).upper() in primary_types]
        trusted_support = [r for r in trusted if clean_text(r.get("Action_Type")).upper() not in primary_types]
        if trusted_primary:
            primary = trusted_primary[0].get("Action_Type")
            if len(trusted_primary) > 1:
                secondary = trusted_primary[1].get("Action_Type")
            elif trusted_support:
                secondary = trusted_support[0].get("Action_Type")
            else:
                secondary = "HISTORICAL_PATTERN"
            behavior = f"{primary}-DRIVEN"
            confidence = "HIGH" if int(p.get("Deal_Bucket_Count", 0) or 0) >= 6 and int(trusted_primary[0].get("Action_Count", 0) or 0) >= 15 else "MEDIUM"
        else:
            primary = "HISTORICAL_PATTERN"
            secondary = trusted_support[0].get("Action_Type") if trusted_support else "NONE"
            behavior = "SPECIFICATION-REPEAT"
            deal_count = int(p.get("Deal_Bucket_Count", 0) or 0)
            confidence = "HIGH" if deal_count >= 6 else "MEDIUM" if deal_count >= 3 else "LOW"
        bulk_count = int(bulk_count_map.get(party, 0))
        if bulk_count >= 2:
            if behavior == "SPECIFICATION-REPEAT":
                behavior = "BULK-OPPORTUNISTIC + SPECIFICATION"
            else:
                behavior = f"{behavior} + BULK-OPPORTUNISTIC"
        conv_map = {r.get("Action_Type"): float(r.get("Party_Action_Conversion_Score", 0.0) or 0.0) for r in ac_rows}
        rows.append({
            "Sold Party": party,
            "Deal_Maker_Class": p.get("Deal_Maker_Class"),
            "Auto_Include": p.get("Auto_Include"),
            "Interval_Status": p.get("Interval_Status"),
            "Deal_Bucket_Count": p.get("Deal_Bucket_Count"),
            "Behavior_Type": behavior,
            "Primary_Trigger": primary,
            "Secondary_Trigger": secondary,
            "Behavior_Confidence": confidence,
            "CART_Conversion_Score": conv_map.get("CART", 0.0),
            "HOLD_Conversion_Score": conv_map.get("HOLD", 0.0),
            "OFFER_Conversion_Score": conv_map.get("OFFER", 0.0),
            "WISHLIST_Conversion_Score": conv_map.get("WISHLIST", 0.0),
            "DETAILS_Conversion_Score": conv_map.get("DETAILS", 0.0),
            "VIEW_Conversion_Score": conv_map.get("VIEW", 0.0),
            "Bulk_Deal_Count": bulk_count,
            "Behavior_Remark": f"Primary trigger: {primary}; model confidence: {confidence}; bulk deal history: {bulk_count}.",
        })
    return pd.DataFrame(rows)

def dedupe_deal_detail_for_profiles(deal_detail: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """Keep one economic stone observation per party/deal/serial.

    A serial may occur on both days of a valid two-day deal when an export
    repeats or revises commercial fields.  Deal recognition already counts
    unique serials; profile denominators and average discount must do the same.
    The latest row in the deal is retained for the final commercial values.
    Blank-serial legacy rows remain separate because they cannot be reconciled.
    """
    if deal_detail is None or deal_detail.empty:
        return pd.DataFrame() if deal_detail is None else deal_detail.copy(), 0
    detail = deal_detail.copy()
    print(f"profile dedupe: rows={len(detail)}", flush=True)
    detail["_Profile_Row_Order"] = np.arange(len(detail), dtype=np.int64)
    with_serial = detail[detail["Serial_Key"].fillna("").astype(str).ne("")].copy()
    no_serial = detail[detail["Serial_Key"].fillna("").astype(str).eq("")].copy()
    before = len(with_serial)
    with_serial = with_serial.sort_values(
        ["Sold Party", "Deal_ID", "Date", "_Profile_Row_Order"], kind="mergesort"
    ).drop_duplicates(["Sold Party", "Deal_ID", "Serial_Key"], keep="last")
    out = pd.concat([with_serial, no_serial], ignore_index=True, sort=False)
    out = out.sort_values(["Sold Party", "Deal_ID", "Date", "_Profile_Row_Order"], kind="mergesort")
    out = out.drop(columns=["_Profile_Row_Order"]).reset_index(drop=True)
    print(f"profile dedupe: canonical_rows={len(out)} removed={before-len(with_serial)}", flush=True)
    return out, before - len(with_serial)


def build_repeated_deal_serial_audit(deal_detail: pd.DataFrame) -> pd.DataFrame:
    """Describe serials repeated inside one historical deal before deduplication.

    The canonical economic/profile row is the latest row in the deal, matching
    ``dedupe_deal_detail_for_profiles``.  This audit preserves the fact that a
    serial appeared on multiple deal dates or rows without allowing those rows
    to inflate stone, carat, amount, discount, or specification calculations.
    """
    columns = [
        "Sold Party", "Deal_ID", "Serial_Key", "Occurrence_Row_Count",
        "Occurrence_Date_Count", "Occurrence_Dates", "First_Occurrence_Date",
        "Last_Occurrence_Date", "Canonical_Kept_Date", "Rows_Removed_From_Profile",
    ]
    if deal_detail is None or deal_detail.empty:
        return pd.DataFrame(columns=columns)
    detail = deal_detail.copy()
    detail["Serial_Key"] = detail.get("Serial_Key", "").fillna("").astype(str)
    detail = detail[detail["Serial_Key"].ne("")].copy()
    if detail.empty:
        return pd.DataFrame(columns=columns)
    detail["Date"] = pd.to_datetime(detail.get("Date"), errors="coerce")
    print(f"repeated audit: rows={len(detail)}", flush=True)
    grouped = detail.groupby(["Sold Party", "Deal_ID", "Serial_Key"], sort=False, dropna=False)
    counts = grouped.size().rename("Occurrence_Row_Count")
    print(f"repeated audit: groups={len(counts)}", flush=True)
    repeated = counts[counts.gt(1)]
    if repeated.empty:
        return pd.DataFrame(columns=columns)
    rows = []
    for key, row_count in repeated.items():
        g = grouped.get_group(key).sort_values("Date", kind="mergesort")
        dates = sorted(pd.to_datetime(g["Date"], errors="coerce").dropna().dt.normalize().unique())
        date_text = " | ".join(pd.Timestamp(d).strftime("%Y-%m-%d") for d in dates)
        first_date = pd.Timestamp(dates[0]) if dates else pd.NaT
        last_date = pd.Timestamp(dates[-1]) if dates else pd.NaT
        rows.append({
            "Sold Party": key[0],
            "Deal_ID": key[1],
            "Serial_Key": key[2],
            "Occurrence_Row_Count": int(row_count),
            "Occurrence_Date_Count": len(dates),
            "Occurrence_Dates": date_text,
            "First_Occurrence_Date": first_date,
            "Last_Occurrence_Date": last_date,
            "Canonical_Kept_Date": last_date,
            "Rows_Removed_From_Profile": int(row_count) - 1,
        })
    return pd.DataFrame(rows, columns=columns).sort_values(
        ["Sold Party", "Deal_ID", "Serial_Key"], kind="mergesort"
    ).reset_index(drop=True)


def build_combination_profiles(
    deal_detail: pd.DataFrame,
    deals: pd.DataFrame,
    settings: dict[str, Any],
) -> pd.DataFrame:
    detail, _ = dedupe_deal_detail_for_profiles(deal_detail)
    if detail.empty:
        return pd.DataFrame()
    detail = enrich_pattern_fields(detail).copy()
    recent_n = int(float(settings.get("RECENT_DEALS_FOR_PATTERN", 5)))
    group_cols = ["Sold Party", "Deal_Mode", "Preferred_Combination_Key"]
    pm_cols = ["Sold Party", "Deal_Mode"]
    dealmeta = deals[["Sold Party", "Deal_ID", "Deal_Start_Date", "Deal_Total_Stones"]].copy()
    dealmeta["Deal_Mode"] = dealmeta["Deal_Total_Stones"].map(mode_for_stones)
    dealmeta = dealmeta.sort_values(["Sold Party", "Deal_Mode", "Deal_Start_Date", "Deal_ID"])
    total_deals = dealmeta.groupby(pm_cols, sort=False)["Deal_ID"].nunique().clip(lower=1)
    recent_deals = dealmeta.groupby(pm_cols, sort=False, group_keys=False).tail(recent_n)
    recent_deal_totals = recent_deals.groupby(pm_cols, sort=False)["Deal_ID"].nunique().clip(lower=1)

    # Party/mode-aware recent membership.  Never use Deal_ID-only membership,
    # even though v6 IDs are globally unique; the composite key is defensive.
    recent_keys = pd.MultiIndex.from_frame(recent_deals[["Sold Party", "Deal_Mode", "Deal_ID"]])
    detail_keys = pd.MultiIndex.from_frame(detail[["Sold Party", "Deal_Mode", "Deal_ID"]])
    detail["_Recent"] = detail_keys.isin(recent_keys)
    detail["_Serial_Nonblank"] = detail["Serial_Key"].where(detail["Serial_Key"].ne(""))
    detail["_Row_Order"] = np.arange(len(detail), dtype=np.int64)
    grp = detail.groupby(group_cols, sort=False, dropna=False)
    size = grp.size().rename("Group_Rows")
    # ``detail`` is already canonicalized to one row per Sold Party + Deal_ID +
    # Serial_Key.  Count rows here so a serial legitimately reused in a later
    # separate deal contributes once to each deal instead of once globally.
    stone_count = size.rename("Stone_Count")
    deal_count = grp["Deal_ID"].nunique().rename("Deal_Count")
    last_date = grp["Date"].max().rename("Last_Sale_Date")
    avg_disc = grp["Disc%"].mean().rename("Avg_Discount")
    first = detail.drop_duplicates(group_cols, keep="first").set_index(group_cols)[["Shp", "Carat_grp", "Color_Band", "Clarity_Band", "Flo_Group", "Lab"]]
    ex = (detail.groupby(group_cols + ["Exact_Combination_Key"], sort=False, dropna=False)
          .agg(_Exact_Count=("Exact_Combination_Key", "size"), _Exact_First=("_Row_Order", "min")).reset_index())
    ex = ex.sort_values(group_cols + ["_Exact_Count", "_Exact_First"], ascending=[True, True, True, False, True])
    rep = ex.drop_duplicates(group_cols).set_index(group_cols)["Exact_Combination_Key"].rename("Representative_Exact_Key")
    r = detail[detail["_Recent"]]
    if not r.empty:
        rg = r.groupby(group_cols, sort=False, dropna=False)
        rsize = rg.size()
        # Same per-deal occurrence rule as Stone_Count; never globally dedupe a
        # serial across separate recent deals.
        recent_count = rsize.rename("Recent_Count")
        recent_deal_count = rg["Deal_ID"].nunique().rename("Recent_Deal_Count")
    else:
        recent_count = pd.Series(dtype=float, name="Recent_Count")
        recent_deal_count = pd.Series(dtype=float, name="Recent_Deal_Count")
    qframes = []
    for dim, col in [("Shade", "Shade_Score"), ("Milky", "Milky_Score"), ("Open", "Open_Score"), ("Blk", "Blk_Score")]:
        vals = pd.to_numeric(detail[col], errors="coerce")
        temp = detail[group_cols].copy()
        temp["_q"] = vals
        temp["_row"] = detail["_Row_Order"]
        valid = temp.dropna(subset=["_q"]).copy()

        def quality_stats(frame: pd.DataFrame, keys: list[str]) -> pd.DataFrame:
            if frame.empty:
                return pd.DataFrame()
            vg = frame.groupby(keys, sort=False, dropna=False)["_q"]
            qmax = vg.max().rename("Exceptional")
            q90 = np.ceil(vg.quantile(.90)).clip(upper=qmax).rename("Normal")
            qc = (frame.groupby(keys + ["_q"], sort=False, dropna=False)
                  .agg(_n=("_q", "size"), _first=("_row", "min")).reset_index())
            qc = qc.sort_values(keys + ["_n", "_q", "_first"], ascending=[True] * len(keys) + [False, True, True])
            qpref = qc.drop_duplicates(keys).set_index(keys)["_q"].rename("Preferred")
            return pd.concat([qpref, q90, qmax], axis=1)

        exact_stats = quality_stats(valid, group_cols)
        pm_stats = quality_stats(valid, pm_cols)
        party_stats = quality_stats(valid, ["Sold Party"])
        known = valid.groupby(group_cols, sort=False, dropna=False).size().reindex(size.index, fill_value=0).astype(int)
        qf = pd.DataFrame(index=size.index)
        qf[f"Known_{dim}_Count"] = known
        qf[f"Unknown_{dim}_Count"] = (size.astype(int) - known).clip(lower=0)
        qf[f"Known_{dim}_Coverage"] = (known / size.clip(lower=1)).astype(float)

        default_name = clean_text(settings.get(f"DEFAULT_HISTORICAL_{dim.upper()}_LIMIT", "NONE")) or "NONE"
        default_score = quality_score(default_name, dim)
        if pd.isna(default_score):
            default_score = quality_score("NONE", dim)
        preferred_values = []
        normal_values = []
        exceptional_values = []
        sources = []
        for key in size.index:
            party, mode, _ = key
            if not exact_stats.empty and key in exact_stats.index:
                stats = exact_stats.loc[key]
                source = "EXACT_COMBINATION_HISTORY"
            elif not pm_stats.empty and (party, mode) in pm_stats.index:
                stats = pm_stats.loc[(party, mode)]
                source = "PARTY_MODE_FALLBACK"
            elif not party_stats.empty and party in party_stats.index:
                stats = party_stats.loc[party]
                source = "PARTY_FALLBACK"
            else:
                stats = pd.Series({"Preferred": default_score, "Normal": default_score, "Exceptional": default_score})
                source = "CONTROL_DEFAULT"
            preferred_values.append(float(stats["Preferred"]))
            normal_values.append(float(stats["Normal"]))
            exceptional_values.append(float(stats["Exceptional"]))
            sources.append(source)
        qf[f"Preferred_{dim}_Score"] = preferred_values
        qf[f"Normal_Max_{dim}_Score"] = normal_values
        qf[f"Exceptional_Max_{dim}_Score"] = exceptional_values
        qf[f"{dim}_Limit_Source"] = sources
        qframes.append(qf)
    out = pd.concat([size, stone_count, deal_count, last_date, avg_disc, first, rep, recent_count, recent_deal_count] + qframes, axis=1).reset_index()
    pm_index = pd.MultiIndex.from_frame(out[pm_cols])
    out["Total_Deals"] = [int(total_deals.get(k, 1)) for k in pm_index]
    total_stones = detail.groupby(pm_cols, sort=False).size().clip(lower=1)
    recent_stones = r.groupby(pm_cols, sort=False).size().clip(lower=1) if not r.empty else pd.Series(dtype=int)
    out["Total_Stones"] = [int(total_stones.get(k, 1)) for k in pm_index]
    out["Recent_Stones"] = [int(recent_stones.get(k, 1)) for k in pm_index]
    out["Recent_Deal_Total"] = [int(recent_deal_totals.get(k, 1)) for k in pm_index]
    out["Recent_Count"] = out["Recent_Count"].fillna(0)
    out["Recent_Deal_Count"] = out["Recent_Deal_Count"].fillna(0)
    out["Deal_Presence_Rate"] = out["Deal_Count"] / out["Total_Deals"]
    out["Stone_Share"] = out["Stone_Count"] / out["Total_Stones"]
    out["Recent_Deal_Presence_Rate"] = out["Recent_Deal_Count"] / out["Recent_Deal_Total"]
    out["Recent_Stone_Share"] = out["Recent_Count"] / out["Recent_Stones"]
    out["Combination_Score"] = (40*out["Deal_Presence_Rate"] + 25*out["Stone_Share"] + 20*out["Recent_Deal_Presence_Rate"] + 15*out["Recent_Stone_Share"]).round(2)
    parts = out["Representative_Exact_Key"].fillna("").str.split("|")
    out["Exact_Col"] = parts.str[2].fillna("UNKNOWN"); out["Exact_Qua"] = parts.str[3].fillna("UNKNOWN"); out["Exact_Flo"] = parts.str[4].fillna("UNKNOWN")
    out["Deal_Mode_Label"] = out["Deal_Mode"].map(MODE_LABEL)
    for dim in ["Shade", "Milky", "Open", "Blk"]:
        out[f"Preferred_{dim}"] = out[f"Preferred_{dim}_Score"].map(lambda x: quality_text(x, dim))
        out[f"Normal_Max_{dim}"] = out[f"Normal_Max_{dim}_Score"].map(lambda x: quality_text(x, dim))
        out[f"Exceptional_Max_{dim}"] = out[f"Exceptional_Max_{dim}_Score"].map(lambda x: quality_text(x, dim))
    cols = ["Sold Party", "Deal_Mode", "Deal_Mode_Label", "Preferred_Combination_Key", "Representative_Exact_Key", "Shp", "Carat_grp", "Color_Band", "Clarity_Band", "Flo_Group", "Lab", "Exact_Col", "Exact_Qua", "Exact_Flo", "Stone_Count", "Deal_Count", "Deal_Presence_Rate", "Stone_Share", "Recent_Deal_Presence_Rate", "Recent_Stone_Share", "Last_Sale_Date", "Avg_Discount", "Combination_Score", "Preferred_Shade", "Normal_Max_Shade", "Exceptional_Max_Shade", "Known_Shade_Count", "Unknown_Shade_Count", "Known_Shade_Coverage", "Shade_Limit_Source", "Preferred_Milky", "Normal_Max_Milky", "Exceptional_Max_Milky", "Known_Milky_Count", "Unknown_Milky_Count", "Known_Milky_Coverage", "Milky_Limit_Source", "Preferred_Open", "Normal_Max_Open", "Exceptional_Max_Open", "Known_Open_Count", "Unknown_Open_Count", "Known_Open_Coverage", "Open_Limit_Source", "Preferred_Blk", "Normal_Max_Blk", "Exceptional_Max_Blk", "Known_Blk_Count", "Unknown_Blk_Count", "Known_Blk_Coverage", "Blk_Limit_Source"]
    return out[cols].sort_values(["Sold Party", "Deal_Mode", "Combination_Score", "Stone_Count"], ascending=[True, True, False, False]).reset_index(drop=True)

def build_recipe_lines(
    combination_profiles: pd.DataFrame,
    mode_summary: pd.DataFrame,
    settings: dict[str, Any],
) -> pd.DataFrame:
    if combination_profiles.empty:
        return pd.DataFrame()
    min_presence = float(settings.get("MIN_RECIPE_DEAL_PRESENCE", 0.15))
    max_lines = int(float(settings.get("MAX_RECIPE_LINES", 15)))
    cumulative_target = float(settings.get("RECIPE_CUMULATIVE_SHARE", 0.90))
    rows = []
    mode_map = mode_summary.set_index(["Sold Party", "Deal_Mode"]).to_dict("index") if not mode_summary.empty else {}
    for (party, mode), g in combination_profiles.groupby(["Sold Party", "Deal_Mode"]):
        m = mode_map.get((party, mode), {})
        target = int(m.get("Target_Stone_Count", 0) or 0)
        work = g[g["Deal_Presence_Rate"] >= min_presence].copy()
        if work.empty:
            work = g.head(max_lines).copy()
        work = work.sort_values(["Combination_Score", "Deal_Presence_Rate", "Stone_Share"], ascending=False)
        work["Cumulative_Share"] = work["Stone_Share"].cumsum()
        selected = work[(work["Cumulative_Share"] <= cumulative_target) | (work.index == work.index[0])].head(max_lines).copy()
        # Include the first row that crosses the target if capacity remains.
        if selected["Stone_Share"].sum() < cumulative_target and len(selected) < max_lines:
            remaining = work.loc[~work.index.isin(selected.index)].head(1)
            selected = pd.concat([selected, remaining], ignore_index=False)
        # A recipe cannot contain more positive-quantity lines than its historical
        # target basket.  This prevents the >=50% presence floor from inflating
        # an 8-stone target into 14 one-stone lines.
        if target <= 0:
            continue
        selected = selected.head(min(len(selected), target)).copy()
        quotas = largest_remainder(selected["Stone_Share"], target)
        selected["Target_Quantity"] = quotas
        # Ensure a line that appeared in >=50% of deals gets at least one slot.
        for idx in selected.index:
            if selected.loc[idx, "Deal_Presence_Rate"] >= 0.50 and selected.loc[idx, "Target_Quantity"] == 0 and target > 0:
                selected.loc[idx, "Target_Quantity"] = 1
        # Rebalance to exact target.
        while int(selected["Target_Quantity"].sum()) > target and target > 0:
            candidates = selected[selected["Target_Quantity"] > 1].sort_values(["Deal_Presence_Rate", "Combination_Score"])
            if candidates.empty:
                break
            selected.loc[candidates.index[0], "Target_Quantity"] -= 1
        while int(selected["Target_Quantity"].sum()) < target and target > 0:
            idx = selected.sort_values(["Combination_Score", "Stone_Share"], ascending=False).index[0]
            selected.loc[idx, "Target_Quantity"] += 1
        selected = selected[selected["Target_Quantity"].gt(0)].copy()
        if int(selected["Target_Quantity"].sum()) != target:
            raise AssertionError(f"Recipe target mismatch for {party} {mode}: target={target}, recipe={int(selected['Target_Quantity'].sum())}")
        for seq, (_, r) in enumerate(selected.iterrows(), 1):
            qty = int(r["Target_Quantity"])
            priority = "CORE" if r["Deal_Presence_Rate"] >= 0.50 else "SUPPORT" if r["Deal_Presence_Rate"] >= 0.25 else "FLEXIBLE"
            rows.append({
                "Sold Party": party,
                "Deal_Mode": mode,
                "Deal_Mode_Label": MODE_LABEL.get(mode),
                "Recipe_Line_ID": f"{re.sub(r'[^A-Z0-9]', '', party.upper())[:8]}-{mode[0]}-{seq:02d}",
                "Recipe_Priority": priority,
                "Preferred_Combination_Key": r["Preferred_Combination_Key"],
                "Representative_Exact_Key": r["Representative_Exact_Key"],
                "Shp": r["Shp"],
                "Carat_grp": r["Carat_grp"],
                "Color_Band": r["Color_Band"],
                "Clarity_Band": r["Clarity_Band"],
                "Flo_Group": r["Flo_Group"],
                "Lab": r["Lab"],
                "Exact_Col": r["Exact_Col"],
                "Exact_Qua": r["Exact_Qua"],
                "Exact_Flo": r["Exact_Flo"],
                "Historical_Target_Stone_Count": target,
                "Target_Quantity": qty,
                "Recipe_Target_Reconciliation_Status": "RECONCILED",
                "Minimum_Quantity": max(qty - 1, 0) if priority != "CORE" else max(math.floor(qty * 0.80), 1 if qty else 0),
                "Maximum_Quantity": qty + 1,
                "Deal_Presence_Rate": r["Deal_Presence_Rate"],
                "Stone_Share": r["Stone_Share"],
                "Combination_Score": r["Combination_Score"],
                "Avg_Discount": r["Avg_Discount"],
                "Preferred_Shade": r["Preferred_Shade"],
                "Normal_Max_Shade": r["Normal_Max_Shade"],
                "Exceptional_Max_Shade": r["Exceptional_Max_Shade"],
                "Known_Shade_Count": int(r.get("Known_Shade_Count", 0) or 0),
                "Unknown_Shade_Count": int(r.get("Unknown_Shade_Count", 0) or 0),
                "Known_Shade_Coverage": float(r.get("Known_Shade_Coverage", 0) or 0),
                "Shade_Limit_Source": r.get("Shade_Limit_Source", ""),
                "Preferred_Milky": r["Preferred_Milky"],
                "Normal_Max_Milky": r["Normal_Max_Milky"],
                "Exceptional_Max_Milky": r["Exceptional_Max_Milky"],
                "Known_Milky_Count": int(r.get("Known_Milky_Count", 0) or 0),
                "Unknown_Milky_Count": int(r.get("Unknown_Milky_Count", 0) or 0),
                "Known_Milky_Coverage": float(r.get("Known_Milky_Coverage", 0) or 0),
                "Milky_Limit_Source": r.get("Milky_Limit_Source", ""),
                "Preferred_Open": r["Preferred_Open"],
                "Normal_Max_Open": r["Normal_Max_Open"],
                "Exceptional_Max_Open": r["Exceptional_Max_Open"],
                "Known_Open_Count": int(r.get("Known_Open_Count", 0) or 0),
                "Unknown_Open_Count": int(r.get("Unknown_Open_Count", 0) or 0),
                "Known_Open_Coverage": float(r.get("Known_Open_Coverage", 0) or 0),
                "Open_Limit_Source": r.get("Open_Limit_Source", ""),
                "Preferred_Blk": r["Preferred_Blk"],
                "Normal_Max_Blk": r["Normal_Max_Blk"],
                "Exceptional_Max_Blk": r["Exceptional_Max_Blk"],
                "Known_Blk_Count": int(r.get("Known_Blk_Count", 0) or 0),
                "Unknown_Blk_Count": int(r.get("Unknown_Blk_Count", 0) or 0),
                "Known_Blk_Coverage": float(r.get("Known_Blk_Coverage", 0) or 0),
                "Blk_Limit_Source": r.get("Blk_Limit_Source", ""),
                "Allow_Relaxed_Match": "YES" if yes(settings.get("ALLOW_RELAXED_MATCH", "YES")) else "NO",
            })
    return pd.DataFrame(rows)


def build_current_actions(
    actions: pd.DataFrame,
    deals: pd.DataFrame,
    action_conversion: pd.DataFrame,
    settings: dict[str, Any],
    as_of: pd.Timestamp,
) -> pd.DataFrame:
    if actions.empty:
        return pd.DataFrame()
    acts = enrich_pattern_fields(actions).copy()
    last_deal = deals.groupby("Sold Party")["Deal_End_Date"].max() if not deals.empty else pd.Series(dtype="datetime64[ns]")
    acts["_Last_Deal"] = acts["Sold Party"].map(last_deal).fillna(pd.Timestamp.min)
    acts = acts[
        acts["Serial_Key"].ne("")
        & acts["Type"].isin(ACTION_TYPES)
        & acts["Date"].gt(acts["_Last_Deal"])
        & acts["Date"].le(as_of)
    ].copy()
    if acts.empty:
        return pd.DataFrame(columns=["Sold Party", "Serial_Key"])
    age_map = {t: int(float(settings.get(ACTION_AGE_SETTING[t], 30))) for t in ACTION_TYPES}
    acts["_Max_Age"] = acts["Type"].map(age_map).fillna(30).astype(int)
    acts["_Age"] = (as_of - acts["Date"]).dt.days
    acts = acts[acts["_Age"].le(acts["_Max_Age"])].copy()
    if acts.empty:
        return pd.DataFrame(columns=["Sold Party", "Serial_Key"])
    key = ["Sold Party", "Serial_Key"]
    type_key = key + ["Type"]
    stats = acts.groupby(type_key, sort=True, dropna=False).agg(
        Type_Action_Count=("Date", "size"),
        Type_First_Date=("Date", "min"),
        Type_Last_Date=("Date", "max"),
        Type_Max_Age=("_Max_Age", "first"),
    ).reset_index()
    if not action_conversion.empty:
        conv = action_conversion[["Sold Party", "Action_Type", "Party_Action_Conversion_Score"]].rename(columns={"Action_Type": "Type"})
        stats = stats.merge(conv, on=["Sold Party", "Type"], how="left")
    else:
        stats["Party_Action_Conversion_Score"] = 0.0
    stats["Party_Action_Conversion_Score"] = pd.to_numeric(stats["Party_Action_Conversion_Score"], errors="coerce").fillna(0.0)
    stats["Type_Age"] = (as_of - stats["Type_Last_Date"]).dt.days.astype(int)
    stats["Recency_Score"] = (1.0 - stats["Type_Age"] / stats["Type_Max_Age"].clip(lower=1)).clip(lower=0.0)
    stats["Repeated_Score"] = stats["Type_Action_Count"].clip(upper=3) / 3.0
    stats["Base_Weight"] = stats["Type"].map(ACTION_BASE_WEIGHTS).fillna(0.25)
    stats["Current_Action_Score"] = (
        stats["Base_Weight"] * (1.0 + stats["Party_Action_Conversion_Score"])
        + stats["Recency_Score"] * 3.0
        + stats["Repeated_Score"] * 2.0
    )
    stats = stats.sort_values(key + ["Current_Action_Score", "Type"], ascending=[True, True, False, True], kind="mergesort")
    primary = stats.drop_duplicates(key, keep="first").copy()
    all_types = stats.groupby(key, sort=False)["Type"].agg(",".join).rename("All_Current_Action_Types")
    overall = acts.groupby(key, sort=False).agg(
        First_Action_Date=("Date", "min"),
        Last_Action_Date=("Date", "max"),
        Action_Count=("Date", "size"),
    )
    latest = acts.sort_values(key + ["Date"], kind="mergesort").drop_duplicates(key, keep="last").set_index(key)
    out = primary.set_index(key).join(all_types).join(overall)
    latest_cols = ["Serial No", "Shp", "Carat", "Carat_grp", "Col", "Qua", "Flo", "Lab", "Preferred_Combination_Key", "Exact_Combination_Key", "Disc%"]
    out = out.join(latest[latest_cols])
    out = out.reset_index()
    out = out.rename(columns={
        "Type": "Primary_Action_Type",
        "Party_Action_Conversion_Score": "Party_Action_Conversion",
        "Type_Age": "Action_Age_Days",
    })
    cols = [
        "Sold Party", "Serial_Key", "Serial No", "Primary_Action_Type", "All_Current_Action_Types",
        "First_Action_Date", "Last_Action_Date", "Action_Age_Days", "Action_Count",
        "Party_Action_Conversion", "Current_Action_Score", "Shp", "Carat", "Carat_grp", "Col", "Qua",
        "Flo", "Lab", "Preferred_Combination_Key", "Exact_Combination_Key", "Disc%",
    ]
    return out[cols].sort_values(["Sold Party", "Current_Action_Score"], ascending=[True, False]).reset_index(drop=True)

# -----------------------------------------------------------------------------
# Stock preparation and exclusions
# -----------------------------------------------------------------------------
def standardize_stock_columns(df: pd.DataFrame) -> pd.DataFrame:
    lookup = {}
    for standard, aliases in STOCK_ALIASES.items():
        for alias in aliases:
            lookup[str(alias).strip().lower()] = standard
    out = df.copy()
    out.columns = [lookup.get(str(c).strip().lower(), str(c).strip()) for c in out.columns]
    return out


def prepare_stock(
    path: Path,
    interval_module,
    settings: dict[str, Any] | None = None,
    quality_log: list[dict[str, Any]] | None = None,
    sheet_name: str | None = None,
    header_row: int | None = None,
    manual_mapping: dict[str, str] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    settings = settings or {}
    quality_log = quality_log if quality_log is not None else []
    required = ["Serial No", "Shp", "Carat", "Col", "Qua", "Flo", "Lab"]
    raw, input_meta = read_dataset_table(
        path, STOCK_ALIASES, required, "Stock",
        sheet_name=sheet_name, header_row=header_row, manual_mapping=manual_mapping,
    )
    stock = standardize_stock_columns(raw)
    status_column_present = "Stock_Status" in stock.columns
    quality_log.extend([
        {"Dataset": "Stock", "Check": "Selected worksheet", "Value": input_meta.sheet_name or "Delimited file", "Severity": "INFO"},
        {"Dataset": "Stock", "Check": "Detected header row", "Value": input_meta.header_row_zero_based + 1, "Severity": "INFO"},
        {"Dataset": "Stock", "Check": "Input columns", "Value": input_meta.column_count, "Severity": "INFO"},
        {"Dataset": "Stock", "Check": "Raw input rows", "Value": len(stock), "Severity": "INFO"},
    ])

    for col in STOCK_ALIASES:
        if col not in stock.columns:
            stock[col] = np.nan
    for col in ["Shade", "Milky", "TOIN_NAME", "SOIN_NAME", "GIRDLE_OPEN_NAME", "PAV_OPEN_NAME", "TBINC_NAME", "SBINC_NAME"]:
        stock[f"Raw_{col}"] = stock[col]
        stock[col] = normalize_upper(stock[col], "")
    for col in ["Shp", "Col", "Qua", "Flo", "Lab", "Cut", "Pol", "Sym"]:
        stock[col] = normalize_upper(stock[col], "")
    stock["Serial_Key"] = serial_key_series(stock["Serial No"])
    stock = annotate_pair_groups(stock)
    stock["Carat"] = pd.to_numeric(stock["Carat"], errors="coerce")
    # Use the interval engine as the single source of truth for stock carat
    # grouping. Recompute whenever numeric Carat is present so old Carat_grp
    # labels do not mix with the revised standard ranges.
    supplied_carat_grp = stock["Carat_grp"].fillna("").astype(str).str.strip().str.upper()
    derived_carat_grp = stock["Carat"].apply(interval_module.get_carat_group)
    stock["Carat_grp"] = derived_carat_grp.where(stock["Carat"].notna(), supplied_carat_grp)
    stock["Disc%"] = pd.to_numeric(stock["Disc%"], errors="coerce")

    stock["Stock_Status"] = stock["Stock_Status"].fillna("").astype(str).str.strip().str.upper()
    stock["Stock_Status_Review_Flag"] = ""
    status_excluded = pd.DataFrame(columns=stock.columns)
    nonblank_status = stock["Stock_Status"].ne("")
    if status_column_present and nonblank_status.any():
        available_values = parse_csv_set(settings.get("STOCK_AVAILABLE_VALUES", "AVAILABLE,IN STOCK,STOCK,YES,Y,1,ACTIVE,OPEN"))
        unavailable_values = parse_csv_set(settings.get("STOCK_UNAVAILABLE_VALUES", "SOLD,RESERVED,UNAVAILABLE,NO,N,0,INACTIVE,MEMO"))
        available_mask = stock["Stock_Status"].isin(available_values)
        unavailable_mask = stock["Stock_Status"].isin(unavailable_values)
        unknown_mask = nonblank_status & ~available_mask & ~unavailable_mask
        policy = clean_text(settings.get("STOCK_UNKNOWN_STATUS_POLICY", "REVIEW")).upper() or "REVIEW"
        if policy not in {"REVIEW", "EXCLUDE", "ERROR", "INCLUDE"}:
            raise ValueError("STOCK_UNKNOWN_STATUS_POLICY must be REVIEW, INCLUDE, EXCLUDE, or ERROR")
        if policy == "ERROR" and unknown_mask.any():
            values = sorted(stock.loc[unknown_mask, "Stock_Status"].unique())
            raise ValueError(f"Unknown stock status values require mapping/review: {values}")
        exclude_mask = unavailable_mask | (unknown_mask if policy == "EXCLUDE" else False)
        stock.loc[unknown_mask, "Stock_Status_Review_Flag"] = "UNKNOWN STATUS - REVIEW"
        status_excluded = stock.loc[exclude_mask].copy()
        if not status_excluded.empty:
            status_excluded["Stock_Status_Review_Flag"] = np.where(
                status_excluded["Stock_Status"].isin(unavailable_values),
                "UNAVAILABLE STATUS",
                "UNKNOWN STATUS EXCLUDED",
            )
        stock = stock.loc[~exclude_mask].copy()
        quality_log.extend([
            {"Dataset": "Stock", "Check": "Availability/status column", "Value": "Stock_Status", "Severity": "OK"},
            {"Dataset": "Stock", "Check": "Rows removed by unavailable status", "Value": int(unavailable_mask.sum()), "Severity": "WARNING" if unavailable_mask.any() else "OK"},
            {"Dataset": "Stock", "Check": "Rows with unknown nonblank status", "Value": int(unknown_mask.sum()), "Severity": "WARNING" if unknown_mask.any() else "OK"},
        ])
    else:
        require_status = yes(settings.get("REQUIRE_STOCK_STATUS_COLUMN", "NO"))
        if require_status:
            raise ValueError("No populated stock availability/status column was detected. Map it in Column_Mapping or set REQUIRE_STOCK_STATUS_COLUMN=NO.")
        stock["Stock_Status_Review_Flag"] = "NO STATUS COLUMN - ASSUMED CURRENT STOCK"
        quality_log.append({
            "Dataset": "Stock",
            "Check": "Availability/status column",
            "Value": "Not detected or blank; all rows treated as current stock",
            "Severity": "WARNING",
        })

    stock["Shade_Score"] = stock["Shade"].apply(lambda x: quality_score(x, "Shade"))
    stock["Milky_Score"] = stock["Milky"].apply(lambda x: quality_score(x, "Milky"))
    stock["Open_Score"] = stock.apply(lambda r: row_max_quality(r, ["TOIN_NAME", "SOIN_NAME", "GIRDLE_OPEN_NAME", "PAV_OPEN_NAME"], "Open"), axis=1)
    stock["Blk_Score"] = stock.apply(lambda r: row_max_quality(r, ["TBINC_NAME", "SBINC_NAME"], "Blk"), axis=1)

    quality_fields = {
        "Shade": ["Shade"], "Milky": ["Milky"],
        "Open": ["TOIN_NAME", "SOIN_NAME", "GIRDLE_OPEN_NAME", "PAV_OPEN_NAME"],
        "Blk": ["TBINC_NAME", "SBINC_NAME"],
    }
    auto_eligible = pd.Series(True, index=stock.index)
    reasons = pd.Series("", index=stock.index, dtype=object)
    for dim, fields in quality_fields.items():
        for field in fields:
            canonical = stock[field].map(lambda x: quality_domain.canonical_quality(x, dim, blank_as_none=False))
            stock[field] = canonical
            eligible = canonical.map(lambda x: quality_domain.quality_auto_eligible(x, dim))
            bad = ~eligible
            auto_eligible &= eligible
            reasons.loc[bad] = reasons.loc[bad].where(reasons.loc[bad].eq(""), reasons.loc[bad] + " | ") + field + ":" + canonical.loc[bad].replace("", "UNKNOWN")
    stock["Quality_Domain_Auto_Eligible"] = auto_eligible.map({True: "YES", False: "NO"})
    stock["Quality_Domain_Review_Flag"] = reasons
    quality_excluded = stock[~auto_eligible].copy()
    auto_exclude = yes(settings.get("QUALITY_REVIEW_VALUES_AUTO_EXCLUDE", "YES")) or clean_text(settings.get("UNKNOWN_QUALITY_POLICY", "EXCLUDE")).upper() == "EXCLUDE"
    if auto_exclude and not quality_excluded.empty:
        stock = stock[auto_eligible].copy()
        status_excluded = concat_nonempty_frames([status_excluded, quality_excluded])
    quality_log.append({"Dataset": "Stock", "Check": "Rows excluded by quality-domain gate", "Value": int(len(quality_excluded)) if auto_exclude else 0, "Severity": "WARNING" if len(quality_excluded) else "OK"})
    stock = enrich_pattern_fields(stock)
    duplicate = stock[stock["Serial_Key"].ne("") & stock.duplicated("Serial_Key", keep=False)].copy()
    stock = stock.drop_duplicates("Serial_Key", keep="first")

    stock, pair_excluded = cascade_incomplete_pair_groups(
        stock,
        "PAIR REMOVED: PARTNER MISSING OR FAILED STOCK/QUALITY ELIGIBILITY",
    )
    if not pair_excluded.empty:
        status_excluded = concat_nonempty_frames([status_excluded, pair_excluded])
    quality_log.append({
        "Dataset": "Stock",
        "Check": "Rows excluded by synchronized pair eligibility",
        "Value": int(len(pair_excluded)),
        "Severity": "WARNING" if len(pair_excluded) else "OK",
    })
    stock = stock.reset_index(drop=True)
    return stock, duplicate, status_excluded


def parse_numeric_condition(series: pd.Series, text: str) -> pd.Series:
    s = pd.to_numeric(series, errors="coerce")
    result = pd.Series(False, index=series.index)
    for raw in str(text).upper().split(","):
        part = raw.strip()
        if not part:
            continue
        try:
            if " TO " in part:
                lo, hi = [float(x.strip()) for x in part.split(" TO ", 1)]
                result |= s.between(lo, hi)
            elif re.match(r"^-?\d+(\.\d+)?\s*-\s*-?\d+(\.\d+)?$", part):
                lo, hi = [float(x.strip()) for x in part.split("-", 1)]
                result |= s.between(lo, hi)
            elif part.startswith(">="):
                result |= s >= float(part[2:].strip())
            elif part.startswith("<="):
                result |= s <= float(part[2:].strip())
            elif part.startswith(">"):
                result |= s > float(part[1:].strip())
            elif part.startswith("<"):
                result |= s < float(part[1:].strip())
            else:
                result |= s.eq(float(part))
        except Exception:
            continue
    return result


def parse_text_condition(series: pd.Series, text: str, order: list[str] | None = None) -> pd.Series:
    s = normalize_upper(series)
    vals = []
    for raw in str(text).upper().split(","):
        part = raw.strip()
        if not part:
            continue
        if " TO " in part and order:
            try:
                a, b = [x.strip() for x in part.split(" TO ", 1)]
                i, j = order.index(a), order.index(b)
                if i > j:
                    i, j = j, i
                vals.extend(order[i:j + 1])
            except Exception:
                vals.append(part)
        else:
            vals.append(part)
    return s.isin(vals)


def apply_exclusions(candidates: pd.DataFrame, rules: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Apply exclusion and keep-only rules.

    Normal values mean "drop rows matching every populated field in this row".
    Prefix a value with ! to make the entire row a keep-only rule. Multiple
    keep-only rows for the same party are OR-ed, while different party/ALL
    groups remain separate controls.
    """
    if candidates.empty or rules.empty or "Sold Party" not in rules.columns:
        return candidates, pd.DataFrame()
    mapping = {
        "SHP": ("Shp", None), "CARAT": ("Carat", "numeric"), "COLOR NAME": ("Col", COLOR_ORDER),
        "QUA": ("Qua", CLARITY_ORDER), "CUT": ("Cut", None), "POL": ("Pol", None), "SYM": ("Sym", None),
        "FLO": ("Flo", FLO_ORDER), "LAB": ("Lab", None), "DISC%": ("Disc%", "numeric"),
        "SHADE": ("Shade", SHADE_ORDER), "MILKY": ("Milky", MILKY_ORDER),
        "TOIN_NAME": ("TOIN_NAME", OPEN_ORDER), "SOIN_NAME": ("SOIN_NAME", OPEN_ORDER),
        "GIRDLE_OPEN_NAME": ("GIRDLE_OPEN_NAME", OPEN_ORDER), "PAV_OPEN_NAME": ("PAV_OPEN_NAME", OPEN_ORDER),
        "TBINC_NAME": ("TBINC_NAME", BLK_ORDER), "SBINC_NAME": ("SBINC_NAME", BLK_ORDER),
    }
    drop = pd.Series(False, index=candidates.index)
    reason = pd.Series("", index=candidates.index, dtype=object)
    keep_groups: dict[str, dict[str, Any]] = {}
    rules2 = rules.copy()
    rules2.columns = [str(c).strip().upper() for c in rules2.columns]
    for ridx, rule in rules2.iterrows():
        party_text = clean_text(rule.get("SOLD PARTY", "ALL")).upper() or "ALL"
        scope = pd.Series(True, index=candidates.index) if party_text == "ALL" else candidates["Sold Party"].astype(str).str.upper().eq(party_text)
        conditions = []
        is_keep_rule = False
        for key, value in rule.items():
            if key in {"SOLD PARTY", "REMARKS"} or pd.isna(value) or clean_text(value) == "":
                continue
            if key not in mapping:
                continue
            text = clean_text(value)
            if text.startswith("!"):
                is_keep_rule = True
                text = text[1:].strip()
            col, order = mapping[key]
            if col not in candidates.columns or not text:
                continue
            if order == "numeric":
                conditions.append(parse_numeric_condition(candidates[col], text))
            else:
                conditions.append(parse_text_condition(candidates[col], text, order))
        if not conditions:
            continue
        match = scope.copy()
        for cond in conditions:
            match &= cond
        remark = clean_text(rule.get("REMARKS"), f"Rule {ridx + 1}")
        if is_keep_rule:
            group = keep_groups.setdefault(party_text, {"scope": scope.copy(), "allowed": pd.Series(False, index=candidates.index), "remarks": []})
            group["allowed"] |= match
            group["remarks"].append(remark)
        else:
            drop |= match
            reason.loc[match & reason.eq("")] = remark
    for party_text, group in keep_groups.items():
        keep_drop = group["scope"] & ~group["allowed"]
        drop |= keep_drop
        keep_reason = "Keep-only rule not matched: " + " | ".join(group["remarks"])
        reason.loc[keep_drop & reason.eq("")] = keep_reason
    removed = candidates.loc[drop].copy()
    if not removed.empty:
        removed["Exclusion_Reason"] = reason.loc[drop]
    return candidates.loc[~drop].copy(), removed


def resolve_party_pair_exclusions(
    stock: pd.DataFrame,
    rules: pd.DataFrame,
    party: str,
) -> tuple[set[str], dict[str, str]]:
    """Return party-specific exclusion serials with pair-wide cascade."""
    if stock is None or stock.empty or rules is None or rules.empty:
        return set(), {}
    universe = stock.copy()
    universe["Sold Party"] = party
    _, direct = apply_exclusions(universe, rules)
    if direct.empty:
        return set(), {}

    direct_serials = set(direct["Serial_Key"].dropna().astype(str))
    direct_reasons = direct.set_index("Serial_Key").get("Exclusion_Reason", pd.Series(dtype=str)).to_dict()
    pair_keys = set(direct.get("Pair_Key", pd.Series(dtype=str)).fillna("").astype(str))
    pair_keys.discard("")
    cascade = universe[universe.get("Pair_Key", pd.Series("", index=universe.index)).fillna("").astype(str).isin(pair_keys)]
    excluded_serials = direct_serials | set(cascade.get("Serial_Key", pd.Series(dtype=str)).dropna().astype(str))

    reasons: dict[str, str] = {}
    for serial in excluded_serials:
        if serial in direct_reasons:
            reasons[serial] = clean_text(direct_reasons[serial], "EXCLUSION RULE")
            continue
        row = universe[universe["Serial_Key"].astype(str).eq(str(serial))]
        key = clean_text(row.iloc[0].get("Pair_Key")) if not row.empty else ""
        source = direct[direct.get("Pair_Key", pd.Series("", index=direct.index)).fillna("").astype(str).eq(key)]
        source_reason = clean_text(source.iloc[0].get("Exclusion_Reason"), "PAIR MEMBER EXCLUDED") if not source.empty else "PAIR MEMBER EXCLUDED"
        reasons[serial] = f"RELATED PAIR MEMBER EXCLUDED: {source_reason}"
    return excluded_serials, reasons


def complete_selected_pairs(
    selected: pd.DataFrame,
    eligible_stock: pd.DataFrame,
    reserved_serials: set[str] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Add every eligible member of a selected pair to the same party basket."""
    if selected is None or selected.empty or "Pair_Key" not in selected.columns:
        return selected.copy(), pd.DataFrame()

    result = selected.copy()
    result["Pair_Added_Flag"] = result.get("Pair_Added_Flag", "NO")
    result["Pair_Anchor_Serial"] = result.get("Pair_Anchor_Serial", "")
    selected_serials = set(result["Serial_Key"].dropna().astype(str))
    selected_pair_keys = set(result["Pair_Key"].fillna("").astype(str))
    selected_pair_keys.discard("")
    if not selected_pair_keys:
        return result, pd.DataFrame()

    additions_source = eligible_stock[
        eligible_stock["Pair_Key"].fillna("").astype(str).isin(selected_pair_keys)
        & ~eligible_stock["Serial_Key"].astype(str).isin(selected_serials)
    ].copy()
    if additions_source.empty:
        return result, pd.DataFrame()

    reserved = set(str(x) for x in (reserved_serials or set()))
    conflicts = set(additions_source["Serial_Key"].astype(str)) & reserved
    if conflicts:
        raise RuntimeError(f"PAIR ALLOCATION CONFLICT: required partner serials already reserved: {sorted(conflicts)}")

    additions: list[pd.Series] = []
    context_columns = [
        "Sold Party", "Deal_Mode", "Recipe_Line_ID", "Recipe_Priority",
        "Preferred_Combination_Key", "Target_Quantity", "Minimum_Quantity",
        "Requested_Recipe_Quantity", "Historical_Target_Quantity",
    ]
    for _, partner in additions_source.iterrows():
        pair_key = clean_text(partner.get("Pair_Key"))
        anchors = result[result["Pair_Key"].fillna("").astype(str).eq(pair_key)]
        if anchors.empty:
            continue
        anchor = anchors.sort_values(
            [c for c in ["Selection_Rank", "Final_Selection_Score"] if c in anchors.columns],
            ascending=[True, False][:len([c for c in ["Selection_Rank", "Final_Selection_Score"] if c in anchors.columns])],
            kind="mergesort",
        ).iloc[0]
        row = partner.copy()
        for col in context_columns:
            if col in anchor.index:
                row[col] = anchor.get(col)
        for col in ["Final_Selection_Score", "Action_Relevance_Score"]:
            if col in anchor.index:
                row[col] = anchor.get(col)
        row["Match_Level"] = "PAIR"
        row["Selection_Source"] = "PAIR COMPLETION"
        row["Selection_Reason"] = f"PAIR PARTNER ADDED FOR {anchor.get('Serial_Key', '')}"
        row["Pair_Added_Flag"] = "YES"
        row["Pair_Anchor_Serial"] = str(anchor.get("Serial_Key", ""))
        row["Direct_Action_Flag"] = row.get("Direct_Action_Flag", "NO") or "NO"
        row["Candidate_Status"] = "PAIR PARTNER ADDED"
        additions.append(row)

    if not additions:
        return result, pd.DataFrame()
    additions_df = pd.DataFrame(additions)
    result = pd.concat([result, additions_df], ignore_index=True, sort=False)
    result = result.drop_duplicates("Serial_Key", keep="first").reset_index(drop=True)
    return result, additions_df


# -----------------------------------------------------------------------------
# Basket selection
# -----------------------------------------------------------------------------
@dataclass
class ModeSelectionResult:
    party: str
    mode: str
    target: int
    selected: pd.DataFrame
    line_status: pd.DataFrame
    candidates: pd.DataFrame
    coverage: float
    core_coverage: float
    accepted: bool
    reason: str


def quality_check(stock: pd.DataFrame, recipe: pd.Series, unknown_policy: str) -> tuple[pd.Series, pd.Series, pd.Series]:
    passed = pd.Series(True, index=stock.index)
    review = pd.Series(False, index=stock.index)
    fit_components = []
    for dim, col in [("Shade", "Shade_Score"), ("Milky", "Milky_Score"), ("Open", "Open_Score"), ("Blk", "Blk_Score")]:
        limit = quality_score(recipe.get(f"Normal_Max_{dim}"), dim)
        preferred = quality_score(recipe.get(f"Preferred_{dim}"), dim)
        vals = pd.to_numeric(stock[col], errors="coerce")
        unknown = vals.isna() | pd.isna(limit)
        review |= unknown
        if unknown_policy == "EXCLUDE":
            passed &= ~unknown
        known = ~unknown
        passed &= (~known) | vals.le(limit)
        if pd.notna(preferred):
            fit = (1 - ((vals - preferred).clip(lower=0) / max(len(QUALITY_ORDERS[dim]) - 1, 1))).clip(lower=0, upper=1)
        else:
            fit = pd.Series(0.5, index=stock.index)
        fit_components.append(fit.fillna(0.25))
    fit_score = sum(fit_components) / len(fit_components) * 10
    return passed, review, fit_score


def recipe_match_level(stock: pd.DataFrame, recipe: pd.Series, allow_relaxed: bool) -> tuple[pd.Series, pd.Series]:
    exact = (
        stock["Shp"].eq(recipe["Shp"])
        & stock["Carat_grp"].eq(recipe["Carat_grp"])
        & stock["Col"].eq(recipe["Exact_Col"])
        & stock["Qua"].eq(recipe["Exact_Qua"])
        & stock["Flo"].eq(recipe["Exact_Flo"])
        & stock["Lab"].eq(recipe["Lab"])
    )
    preferred = (
        stock["Shp"].eq(recipe["Shp"])
        & stock["Carat_grp"].eq(recipe["Carat_grp"])
        & stock["Color_Band"].eq(recipe["Color_Band"])
        & stock["Clarity_Band"].eq(recipe["Clarity_Band"])
        & stock["Flo_Group"].eq(recipe["Flo_Group"])
        & stock["Lab"].eq(recipe["Lab"])
    )
    relaxed = (
        stock["Shp"].eq(recipe["Shp"])
        & stock["Carat_grp"].eq(recipe["Carat_grp"])
        & stock["Clarity_Band"].eq(recipe["Clarity_Band"])
        & stock["Lab"].eq(recipe["Lab"])
    ) if allow_relaxed else pd.Series(False, index=stock.index)
    level = pd.Series("NO MATCH", index=stock.index, dtype=object)
    base = pd.Series(0.0, index=stock.index)
    level.loc[relaxed] = "RELAXED"; base.loc[relaxed] = 55
    level.loc[preferred] = "PREFERRED"; base.loc[preferred] = 80
    level.loc[exact] = "EXACT"; base.loc[exact] = 100
    return level, base


def score_recipe_candidates(
    party: str,
    mode: str,
    recipe: pd.Series,
    stock: pd.DataFrame,
    current_actions: pd.DataFrame,
    behavior: pd.Series,
    settings: dict[str, Any],
) -> pd.DataFrame:
    allow_relaxed = yes(recipe.get("Allow_Relaxed_Match", settings.get("ALLOW_RELAXED_MATCH", "YES")))
    level, match_score = recipe_match_level(stock, recipe, allow_relaxed)
    mask = level.ne("NO MATCH")
    candidates = stock.loc[mask].copy()
    if candidates.empty:
        return candidates
    candidates["Match_Level"] = level.loc[mask]
    candidates["Match_Base_Score"] = match_score.loc[mask]
    unknown_policy = clean_text(settings.get("UNKNOWN_QUALITY_POLICY", "REVIEW")).upper()
    qpass, qreview, qfit = quality_check(candidates, recipe, unknown_policy)
    candidates = candidates.loc[qpass].copy()
    if candidates.empty:
        return candidates
    candidates["Quality_Review_Flag"] = qreview.loc[candidates.index].map({True: "YES", False: "NO"})
    candidates["Quality_Fit_Score"] = qfit.loc[candidates.index]
    actions = current_actions[current_actions["Sold Party"].eq(party)].copy() if not current_actions.empty else pd.DataFrame()
    if not actions.empty:
        action_cols = ["Serial_Key", "Primary_Action_Type", "Last_Action_Date", "Action_Age_Days", "Action_Count", "Party_Action_Conversion", "Current_Action_Score"]
        candidates = candidates.merge(actions[action_cols], on="Serial_Key", how="left")
    else:
        for col in ["Primary_Action_Type", "Last_Action_Date", "Action_Age_Days", "Action_Count", "Party_Action_Conversion", "Current_Action_Score"]:
            candidates[col] = np.nan
    candidates["Current_Action_Score"] = pd.to_numeric(candidates["Current_Action_Score"], errors="coerce").fillna(0)
    candidates["Party_Action_Conversion"] = pd.to_numeric(candidates["Party_Action_Conversion"], errors="coerce").fillna(0)
    primary_trigger = clean_text(behavior.get("Primary_Trigger", "HISTORICAL_PATTERN")).upper()
    trigger_bonus = candidates["Primary_Action_Type"].astype(str).str.upper().eq(primary_trigger).astype(float) * 10
    action_score = (candidates["Current_Action_Score"].clip(upper=15) / 15 * 25) + trigger_bonus
    candidates["Action_Relevance_Score"] = action_score
    historical_score = float(recipe.get("Combination_Score", 0)) / 100 * 30
    candidates["Historical_Combination_Score"] = historical_score
    avg_disc = pd.to_numeric(pd.Series([recipe.get("Avg_Discount")]), errors="coerce").iloc[0]
    if pd.notna(avg_disc):
        diff = (pd.to_numeric(candidates["Disc%"], errors="coerce") - avg_disc).abs()
        candidates["Discount_Fit_Score"] = (10 - diff / 2).clip(lower=0, upper=10).fillna(3)
    else:
        candidates["Discount_Fit_Score"] = 6.0
    candidates["Final_Selection_Score"] = (
        candidates["Match_Base_Score"]
        + candidates["Historical_Combination_Score"]
        + candidates["Action_Relevance_Score"]
        + candidates["Quality_Fit_Score"]
        + candidates["Discount_Fit_Score"]
    )
    candidates["Sold Party"] = party
    candidates["Deal_Mode"] = mode
    candidates["Recipe_Line_ID"] = recipe["Recipe_Line_ID"]
    candidates["Recipe_Priority"] = recipe["Recipe_Priority"]
    candidates["Target_Quantity"] = recipe["Target_Quantity"]
    direct_types = parse_csv_set(settings.get("DIRECT_SELECTION_ACTION_TYPES", "HOLD,CART,OFFER"))
    primary_action_upper = candidates["Primary_Action_Type"].fillna("").astype(str).str.upper()
    direct_mask = primary_action_upper.isin(direct_types)
    candidates["Direct_Action_Flag"] = np.where(direct_mask, "YES", "NO")
    candidates["Selection_Source"] = np.where(
        direct_mask,
        "CURRENT_" + primary_action_upper,
        "HISTORICAL_RECIPE",
    )
    candidates["Selection_Reason"] = candidates.apply(
        lambda r: f"{r['Match_Level']} recipe fit; {recipe['Deal_Presence_Rate']:.0%} historical deal presence"
        + (
            f"; direct current {r['Primary_Action_Type']}"
            if clean_text(r.get("Primary_Action_Type")).upper() in direct_types
            else (f"; supporting current {r['Primary_Action_Type']}" if pd.notna(r.get("Primary_Action_Type")) else "")
        ),
        axis=1,
    )
    policy = clean_text(settings.get("DIRECT_ACTION_POLICY", "DIRECT_FIRST_WITHIN_MATCH")).upper()
    if policy == "DIRECT_FIRST_WITHIN_MATCH":
        candidates["_Match_Rank"] = candidates["Match_Level"].map({"EXACT": 0, "PREFERRED": 1, "RELAXED": 2}).fillna(9)
        candidates["_Direct_Rank"] = np.where(candidates["Direct_Action_Flag"].eq("YES"), 0, 1)
        candidates = candidates.sort_values(["_Match_Rank", "_Direct_Rank", "Final_Selection_Score", "Action_Relevance_Score"], ascending=[True, True, False, False])
        return candidates.drop(columns=["_Match_Rank", "_Direct_Rank"])
    if policy != "SCORE_WEIGHTED":
        raise ValueError("DIRECT_ACTION_POLICY must be DIRECT_FIRST_WITHIN_MATCH or SCORE_WEIGHTED")
    return candidates.sort_values(["Final_Selection_Score", "Action_Relevance_Score"], ascending=False)


def select_for_mode(
    party: str,
    mode: str,
    recipes: pd.DataFrame,
    stock: pd.DataFrame,
    current_actions: pd.DataFrame,
    behavior: pd.Series,
    settings: dict[str, Any],
    exclusions: pd.DataFrame,
    stock_index: dict[tuple[str, str, str, str], pd.DataFrame] | None = None,
    party_actions: pd.DataFrame | None = None,
    reserved_serials: set[str] | None = None,
    reserved_pair_keys: set[str] | None = None,
) -> ModeSelectionResult:
    rec = recipes[(recipes["Sold Party"].eq(party)) & (recipes["Deal_Mode"].eq(mode))].copy()
    target = int(rec["Target_Quantity"].sum()) if not rec.empty else 0
    if rec.empty or target <= 0:
        return ModeSelectionResult(party, mode, target, pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), 0, 0, False, "No recipe lines")
    used: set[str] = set()
    selected_parts = []
    candidate_parts = []
    line_rows = []
    priority_order = {"CORE": 0, "SUPPORT": 1, "FLEXIBLE": 2}
    rec["_priority"] = rec["Recipe_Priority"].map(priority_order).fillna(9)
    rec = rec.sort_values(["_priority", "Combination_Score"], ascending=[True, False])
    empty_stock = stock.iloc[0:0]
    actions_for_party = party_actions if party_actions is not None else current_actions
    pair_excluded_serials, pair_exclusion_reasons = resolve_party_pair_exclusions(stock, exclusions, party)
    for _, line in rec.iterrows():
        if stock_index is not None:
            stock_key = (clean_text(line.get("Shp")).upper(), clean_text(line.get("Carat_grp")), clean_text(line.get("Lab")).upper(), clean_text(line.get("Clarity_Band")).upper())
            line_stock = stock_index.get(stock_key, empty_stock)
        else:
            line_stock = stock
        if reserved_serials:
            line_stock = line_stock[~line_stock["Serial_Key"].isin(reserved_serials)]
        if reserved_pair_keys and "Pair_Key" in line_stock.columns:
            line_stock = line_stock[~line_stock["Pair_Key"].fillna("").astype(str).isin(reserved_pair_keys)]
        scored = score_recipe_candidates(party, mode, line, line_stock, actions_for_party, behavior, settings)
        if not scored.empty and pair_excluded_serials:
            removed_mask = scored["Serial_Key"].astype(str).isin(pair_excluded_serials)
            removed = scored.loc[removed_mask].copy()
            if not removed.empty:
                removed["Exclusion_Reason"] = removed["Serial_Key"].astype(str).map(pair_exclusion_reasons).fillna("PAIR-SYNCHRONIZED EXCLUSION")
                removed["Candidate_Status"] = "EXCLUDED"
                candidate_parts.append(removed)
            scored = scored.loc[~removed_mask].copy()
        scored = scored[~scored["Serial_Key"].isin(used)].copy() if not scored.empty else scored
        qty = int(line["Target_Quantity"])
        chosen = scored.head(qty).copy() if not scored.empty else pd.DataFrame()
        if not chosen.empty:
            used.update(chosen["Serial_Key"].tolist())
            selected_parts.append(chosen)
        if not scored.empty:
            scored2 = scored.copy()
            scored2["Candidate_Status"] = np.where(scored2["Serial_Key"].isin(set(chosen.get("Serial_Key", []))), "SELECTED", "NOT SELECTED")
            candidate_parts.append(scored2)
        line_rows.append({
            "Sold Party": party,
            "Deal_Mode": mode,
            "Recipe_Line_ID": line["Recipe_Line_ID"],
            "Recipe_Priority": line["Recipe_Priority"],
            "Preferred_Combination_Key": line["Preferred_Combination_Key"],
            "Target_Quantity": qty,
            "Requested_Recipe_Quantity": qty,
            "Minimum_Quantity": line["Minimum_Quantity"],
            "Available_Qualifying_Stock": len(scored),
            "Candidate_Availability": len(scored),
            "Selected_Quantity": len(chosen),
            "Pre_Allocation_Filled_Quantity": len(chosen),
            "Line_Coverage": len(chosen) / qty if qty else 1.0,
        })
    selected = pd.concat(selected_parts, ignore_index=True) if selected_parts else pd.DataFrame()
    candidates = pd.concat(candidate_parts, ignore_index=True) if candidate_parts else pd.DataFrame()
    line_status = pd.DataFrame(line_rows)

    base_selected = selected.copy()
    base_line_counts = (
        base_selected.groupby("Recipe_Line_ID").size()
        if not base_selected.empty and "Recipe_Line_ID" in base_selected.columns
        else pd.Series(dtype=int)
    )
    base_coverage = len(base_selected) / target if target else 0
    core = line_status[line_status["Recipe_Priority"].eq("CORE")]
    core_target = core["Target_Quantity"].sum()
    base_core_selected = core["Recipe_Line_ID"].map(base_line_counts).fillna(0).sum() if not core.empty else 0
    base_core_coverage = base_core_selected / core_target if core_target else 1.0

    selected, pair_additions = complete_selected_pairs(base_selected, stock, reserved_serials)
    pair_added_count = len(pair_additions)
    if not pair_additions.empty:
        pair_candidates = pair_additions.copy()
        pair_candidates["Candidate_Status"] = "PAIR PARTNER ADDED"
        candidates = pd.concat([candidates, pair_candidates], ignore_index=True, sort=False) if not candidates.empty else pair_candidates

    if not line_status.empty:
        final_counts = (
            selected.groupby("Recipe_Line_ID").size()
            if not selected.empty and "Recipe_Line_ID" in selected.columns
            else pd.Series(dtype=int)
        )
        normal_counts = line_status["Recipe_Line_ID"].map(base_line_counts).fillna(0).astype(int)
        total_counts = line_status["Recipe_Line_ID"].map(final_counts).fillna(0).astype(int)
        line_status["Normal_Selected_Quantity"] = normal_counts
        line_status["Pair_Added_Quantity"] = (total_counts - normal_counts).clip(lower=0)
        line_status["Selected_Quantity"] = total_counts
        line_status["Pre_Allocation_Filled_Quantity"] = total_counts
        requested_qty = pd.to_numeric(line_status["Target_Quantity"], errors="coerce").fillna(0)
        line_status["Line_Coverage"] = np.where(requested_qty.gt(0), normal_counts / requested_qty, 1.0)

    min_coverage = float(settings.get(MODE_COVERAGE_SETTING[mode], 0.70))
    min_core = float(settings.get("CORE_MIN_STOCK_COVERAGE", 0.90))
    absolute_min = MODE_MIN_STONES.get(mode, 0) if yes(settings.get("ENFORCE_ABSOLUTE_MODE_MINIMUM", "YES")) else 0
    absolute_max = MODE_MAX_STONES.get(mode)
    base_size_ok = len(base_selected) >= absolute_min and (absolute_max is None or len(base_selected) <= absolute_max)
    accepted = base_coverage >= min_coverage and base_core_coverage >= min_core and base_size_ok
    size_text = f"normal size {len(base_selected)} (required {absolute_min}+" + (f", maximum {absolute_max}" if absolute_max is not None else "") + ")"
    reason = f"Coverage {base_coverage:.0%} (required {min_coverage:.0%}); core {base_core_coverage:.0%} (required {min_core:.0%}); {size_text}"
    if pair_added_count:
        reason += f"; pair completion added {pair_added_count} partner stone(s); final basket {len(selected)}"
    if not selected.empty:
        sort_cols = [c for c in ["Final_Selection_Score", "Pair_Added_Flag"] if c in selected.columns]
        if sort_cols:
            selected = selected.sort_values(sort_cols, ascending=[False, True][:len(sort_cols)], kind="mergesort").reset_index(drop=True)
        selected["Selection_Rank"] = np.arange(1, len(selected) + 1)
        selected["Pair_Review_Flag"] = np.where(selected["Pair_Key"].fillna("").astype(str).ne(""), "YES", "NO")
    return ModeSelectionResult(party, mode, target, selected, line_status, candidates, base_coverage, base_core_coverage, accepted, reason)



def reconcile_recipe_fill(
    line_status: pd.DataFrame,
    selected: pd.DataFrame,
    party_status: pd.DataFrame,
) -> pd.DataFrame:
    """Reconcile attempted recipe fills to the final globally allocated basket."""
    if line_status is None or line_status.empty:
        return pd.DataFrame() if line_status is None else line_status.copy()
    out = line_status.copy()
    requested = pd.to_numeric(
        out.get("Requested_Recipe_Quantity", out.get("Target_Quantity", 0)), errors="coerce"
    ).fillna(0).astype(int)
    availability = pd.to_numeric(
        out.get("Candidate_Availability", out.get("Available_Qualifying_Stock", 0)), errors="coerce"
    ).fillna(0).astype(int)
    prefill = pd.to_numeric(
        out.get("Pre_Allocation_Filled_Quantity", out.get("Selected_Quantity", 0)), errors="coerce"
    ).fillna(0).astype(int)
    out["Requested_Recipe_Quantity"] = requested
    out["Candidate_Availability"] = availability
    out["Pre_Allocation_Filled_Quantity"] = prefill

    keys = ["Sold Party", "Deal_Mode", "Recipe_Line_ID"]
    if selected is not None and not selected.empty and set(keys).issubset(selected.columns):
        final_counts = selected.groupby(keys, dropna=False).size().rename("Final_Globally_Allocated_Quantity").reset_index()
        out = out.merge(final_counts, on=keys, how="left")
    else:
        out["Final_Globally_Allocated_Quantity"] = 0
    out["Final_Globally_Allocated_Quantity"] = pd.to_numeric(
        out["Final_Globally_Allocated_Quantity"], errors="coerce"
    ).fillna(0).astype(int)
    out["Missing_Quantity"] = (
        out["Requested_Recipe_Quantity"] - out["Final_Globally_Allocated_Quantity"]
    ).clip(lower=0).astype(int)
    out["Final_Line_Coverage"] = np.where(
        out["Requested_Recipe_Quantity"].gt(0),
        out["Final_Globally_Allocated_Quantity"] / out["Requested_Recipe_Quantity"],
        1.0,
    )

    status_map = {}
    if party_status is not None and not party_status.empty and {"Sold Party", "Generation_Status"}.issubset(party_status.columns):
        status_map = party_status.drop_duplicates("Sold Party", keep="last").set_index("Sold Party")["Generation_Status"].to_dict()
    out["Party_Generation_Status"] = out["Sold Party"].map(status_map).fillna("UNKNOWN")

    def reason(row: pd.Series) -> str:
        chosen = clean_text(row.get("Chosen_Mode", "NO")).upper() == "YES"
        generated = clean_text(row.get("Party_Generation_Status", "")).upper() == "GENERATED"
        requested_qty = int(row.get("Requested_Recipe_Quantity", 0) or 0)
        available_qty = int(row.get("Candidate_Availability", 0) or 0)
        pre_qty = int(row.get("Pre_Allocation_Filled_Quantity", 0) or 0)
        final_qty = int(row.get("Final_Globally_Allocated_Quantity", 0) or 0)
        if not chosen:
            return "MODE NOT CHOSEN"
        if not generated:
            if available_qty <= 0:
                return "NO QUALIFYING CANDIDATES; BASKET NOT ACCEPTED"
            if pre_qty < requested_qty:
                return "INSUFFICIENT QUALIFYING STOCK; PARTIAL ATTEMPT NOT ALLOCATED"
            return "BASKET FAILED COVERAGE/CORE GATE; PRESELECTION NOT ALLOCATED"
        if final_qty >= requested_qty:
            return "FILLED"
        if final_qty < pre_qty:
            return "FINAL GLOBAL ALLOCATION REDUCTION"
        if available_qty <= 0:
            return "NO QUALIFYING CANDIDATES"
        if available_qty < requested_qty or pre_qty < requested_qty:
            return "INSUFFICIENT QUALIFYING STOCK"
        return "PARTIAL FINAL FILL"

    out["Failure_Reason"] = out.apply(reason, axis=1)
    chosen = out.get("Chosen_Mode", pd.Series("NO", index=out.index)).fillna("NO").astype(str).str.upper().eq("YES")
    generated = out["Party_Generation_Status"].astype(str).str.upper().eq("GENERATED")
    expected_final = np.where(chosen & generated, out["Pre_Allocation_Filled_Quantity"], 0)
    out["Recipe_Fill_Reconciliation_Status"] = np.where(
        out["Final_Globally_Allocated_Quantity"].to_numpy() == expected_final,
        "RECONCILED",
        "MISMATCH",
    )
    return out


def annotate_candidate_archive(
    candidates: pd.DataFrame,
    selected: pd.DataFrame,
    party_status: pd.DataFrame,
) -> pd.DataFrame:
    """Make candidate archive stage and final-allocation semantics explicit."""
    if candidates is None or candidates.empty:
        return pd.DataFrame() if candidates is None else candidates.copy()
    out = candidates.copy()
    out["Pre_Allocation_Candidate_Status"] = out.get("Candidate_Status", "")
    out["Candidate_Archive_Stage"] = (
        "RANKED ATTEMPT CANDIDATES - POST QUALITY/EXCLUSIONS/PRIOR GLOBAL "
        "SERIAL-PAIR RESERVATIONS; PRE FINAL BASKET ACCEPTANCE"
    )
    out["Candidate_Archive_Is_Final_Output"] = "NO"
    keys = ["Sold Party", "Deal_Mode", "Recipe_Line_ID", "Serial_Key"]
    final_keys: set[tuple[str, str, str, str]] = set()
    if selected is not None and not selected.empty and set(keys).issubset(selected.columns):
        final_keys = set(map(tuple, selected[keys].fillna("").astype(str).itertuples(index=False, name=None)))
    row_keys = list(map(tuple, out.reindex(columns=keys, fill_value="").fillna("").astype(str).itertuples(index=False, name=None)))
    out["Final_Globally_Allocated_Flag"] = ["YES" if key in final_keys else "NO" for key in row_keys]

    status_map = {}
    if party_status is not None and not party_status.empty and {"Sold Party", "Generation_Status"}.issubset(party_status.columns):
        status_map = party_status.drop_duplicates("Sold Party", keep="last").set_index("Sold Party")["Generation_Status"].to_dict()
    out["Party_Generation_Status"] = out.get("Sold Party", pd.Series(index=out.index, dtype=str)).map(status_map).fillna("UNKNOWN")

    def allocation_status(row: pd.Series) -> str:
        if clean_text(row.get("Final_Globally_Allocated_Flag")).upper() == "YES":
            return "FINAL ALLOCATED"
        pre = clean_text(row.get("Pre_Allocation_Candidate_Status")).upper()
        if pre == "EXCLUDED":
            return "EXCLUDED BY BUSINESS RULE"
        if clean_text(row.get("Chosen_Mode", "NO")).upper() != "YES":
            return "MODE NOT CHOSEN"
        if clean_text(row.get("Party_Generation_Status")).upper() != "GENERATED":
            return "PARTIAL ATTEMPT - BASKET NOT ACCEPTED"
        if pre == "SELECTED":
            return "PRESELECTED BUT NOT FINAL - REVIEW"
        return "RANKED NOT SELECTED"

    out["Final_Allocation_Status"] = out.apply(allocation_status, axis=1)
    return out


def determine_candidate_modes(
    party: str,
    party_row: pd.Series,
    behavior: pd.Series,
    mode_summary: pd.DataFrame,
    overrides: pd.DataFrame,
    settings: dict[str, Any],
) -> list[str]:
    force_mode = "AUTO"
    if not overrides.empty and "Sold Party" in overrides.columns:
        ov = overrides[overrides["Sold Party"].fillna("").astype(str).str.strip().eq(party)]
        if not ov.empty:
            force_mode = clean_text(ov.iloc[-1].get("Force_Deal_Mode", "AUTO")).upper() or "AUTO"
    if force_mode in MODE_ORDER:
        return [force_mode]
    pm = mode_summary[mode_summary["Sold Party"].eq(party)].copy()
    if pm.empty:
        return []
    include_due_soon = yes(settings.get("INCLUDE_DUE_SOON", "NO"))
    allowed = {"DUE NOW", "OVERDUE"} | ({"DUE SOON"} if include_due_soon else set())
    due_modes = pm[pm["Mode_Interval_Status"].isin(allowed)].copy()
    if yes(settings.get("MODE_DUE_REQUIRED_FOR_SELECTION", "YES")):
        candidate_modes = due_modes["Deal_Mode"].tolist()
    else:
        candidate_modes = due_modes["Deal_Mode"].tolist()
        if clean_text(party_row.get("Auto_Include")).upper() == "YES" and yes(settings.get("ALLOW_CROSS_MODE_FALLBACK", "NO")):
            all_modes = pm.sort_values(["Mode_Deal_Count", "Share_of_Party_Deals"], ascending=False)["Deal_Mode"].tolist()
            candidate_modes += [m for m in all_modes if m not in candidate_modes]
    if "BULK-OPPORTUNISTIC" in clean_text(behavior.get("Behavior_Type")).upper():
        order = {"BULK": 0, "MEDIUM": 1, "SMALL": 2}
    else:
        order = {"SMALL": 0, "MEDIUM": 1, "BULK": 2}
    return sorted(set(candidate_modes), key=lambda x: order.get(x, 9))

def action_policy(settings: dict[str, Any]) -> str:
    policy = clean_text(settings.get("ACTION_POLICY", "BOOST_ONLY")).upper() or "BOOST_ONLY"
    if policy not in {"BOOST_ONLY", "HARD_GATE"}:
        raise ValueError("ACTION_POLICY must be BOOST_ONLY or HARD_GATE")
    return policy


def action_support_threshold(mode: str, settings: dict[str, Any]) -> int:
    mode = clean_text(mode).upper()
    if mode == "BULK":
        return int(float(settings.get("BULK_MIN_CURRENT_ACTION_STONES", 11)))
    if mode == "MEDIUM":
        return int(float(settings.get("MEDIUM_MIN_CURRENT_ACTION_STONES", 5)))
    return 1


def classify_action_support(mode: str, current_action_stones: int, settings: dict[str, Any], actions_staleness_days: int | None = None) -> dict[str, Any]:
    threshold = action_support_threshold(mode, settings)
    supported = int(current_action_stones) >= threshold
    max_age = int(float(settings.get("MAX_ACTIONS_STALENESS_DAYS", 7)))
    stale = actions_staleness_days is None or int(actions_staleness_days) > max_age
    data_status = "STALE ACTION DATA" if stale else ("NO RECENT ACTION DATA" if int(current_action_stones) == 0 else "CURRENT ACTION DATA")
    support_status = "ACTION SUPPORTED" if supported else "SALES HISTORY ONLY"
    review_required = stale or not supported
    return {
        "Action_Policy": action_policy(settings),
        "Action_Support_Threshold": threshold,
        "Action_Support_Passed": "YES" if supported else "NO",
        "Action_Support_Status": support_status,
        "Action_Data_Status": data_status,
        "Action_Review_Required": "YES" if review_required else "NO",
    }


def choose_mode_result(results: list[ModeSelectionResult], current_actions: pd.DataFrame, party: str, settings: dict[str, Any]) -> ModeSelectionResult | None:
    if not results:
        return None
    available_action_stones = current_actions[current_actions["Sold Party"].eq(party)]["Serial_Key"].nunique() if not current_actions.empty else 0
    accepted = [r for r in results if r.accepted]
    if not accepted:
        # Actions cannot override insufficient stock, quality, pair, recipe, or size controls.
        return max(results, key=lambda r: (r.coverage, r.core_coverage, -MODE_ORDER.index(r.mode)))
    if action_policy(settings) == "BOOST_ONLY":
        # Completed Sales history determines which modes are candidates. Actions already
        # influence ranking inside each historical match level, but never reject a valid basket.
        return max(accepted, key=lambda r: (MODE_ORDER.index(r.mode), r.coverage, r.core_coverage))
    bulk_min = int(float(settings.get("BULK_MIN_CURRENT_ACTION_STONES", 11)))
    medium_min = int(float(settings.get("MEDIUM_MIN_CURRENT_ACTION_STONES", 5)))
    eligible = []
    blocked = []
    for r in accepted:
        if r.mode == "BULK" and yes(settings.get("BULK_ACTION_REQUIREMENT_IS_HARD_GATE", "YES")) and available_action_stones < bulk_min:
            blocked.append((r, f"BULK action gate not met: {available_action_stones} < {bulk_min}"))
            continue
        if r.mode == "MEDIUM" and yes(settings.get("MEDIUM_ACTION_REQUIREMENT_IS_HARD_GATE", "YES")) and available_action_stones < medium_min:
            blocked.append((r, f"MEDIUM action gate not met: {available_action_stones} < {medium_min}"))
            continue
        eligible.append(r)
    if eligible:
        return max(eligible, key=lambda r: (MODE_ORDER.index(r.mode), r.coverage, r.core_coverage))
    base, gate_reason = max(blocked, key=lambda item: (item[0].coverage, item[0].core_coverage, -MODE_ORDER.index(item[0].mode)))
    return ModeSelectionResult(base.party, base.mode, base.target, base.selected, base.line_status, base.candidates,
                               base.coverage, base.core_coverage, False, f"{base.reason}; {gate_reason}")

def generate_deals(
    party_summary: pd.DataFrame,
    behavior_profiles: pd.DataFrame,
    mode_summary: pd.DataFrame,
    recipes: pd.DataFrame,
    current_actions: pd.DataFrame,
    stock: pd.DataFrame,
    overrides: pd.DataFrame,
    removed: pd.DataFrame,
    exclusions: pd.DataFrame,
    settings: dict[str, Any],
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    removed_set = set(removed.get("Sold Party", pd.Series(dtype=str)).dropna().astype(str).str.strip()) if not removed.empty else set()
    status_rows = []
    selected_parts = []
    line_parts = []
    candidate_parts = []
    bmap = behavior_profiles.set_index("Sold Party").to_dict("index") if not behavior_profiles.empty else {}
    stock_index = {
        (clean_text(k[0]).upper(), clean_text(k[1]), clean_text(k[2]).upper(), clean_text(k[3]).upper()): g
        for k, g in stock.groupby(["Shp", "Carat_grp", "Lab", "Clarity_Band"], sort=False, dropna=False)
    } if not stock.empty else {}
    action_index = {party: g for party, g in current_actions.groupby("Sold Party", sort=False)} if not current_actions.empty else {}
    allocation_mode = clean_text(settings.get("GLOBAL_STOCK_ALLOCATION", "INDEPENDENT")).upper()
    unique_allocation = allocation_mode == "UNIQUE"
    global_used: set[str] = set()
    party_work = party_summary.copy()
    if unique_allocation:
        party_work["_Eligible_Order"] = party_work["Auto_Include"].astype(str).str.upper().eq("YES").astype(int)
        party_work = party_work.sort_values(
            ["_Eligible_Order", "Interval_Progress_Ratio", "Recent_Action_Count", "Deal_Bucket_Count", "Sold Party"],
            ascending=[False, False, False, False, True],
            kind="mergesort",
        )
    for _, p in party_work.iterrows():
        party = p["Sold Party"]
        if party in removed_set:
            status_rows.append({"Sold Party": party, "Generation_Status": "EXCLUDED", "Generation_Reason": "Removed party control"})
            continue
        if yes(settings.get("OUTPUT_ONLY_INTERVAL_ELIGIBLE", "YES")) and clean_text(p.get("Auto_Include")).upper() != "YES":
            status_rows.append({
                "Sold Party": party,
                "Generation_Status": "NOT DUE",
                "Generation_Reason": p.get("Auto_Include_Reason"),
                "Interval_Status": p.get("Interval_Status"),
            })
            continue
        behavior = pd.Series(bmap.get(party, {"Primary_Trigger": "HISTORICAL_PATTERN", "Behavior_Type": "SPECIFICATION-REPEAT"}))
        modes = determine_candidate_modes(party, p, behavior, mode_summary, overrides, settings)
        party_actions = action_index.get(party, current_actions.iloc[0:0] if not current_actions.empty else pd.DataFrame())
        results = [
            select_for_mode(
                party, m, recipes, stock, current_actions, behavior, settings, exclusions,
                stock_index, party_actions, global_used if unique_allocation else None,
            )
            for m in modes
        ]
        chosen = choose_mode_result(results, current_actions, party, settings)
        if chosen is None:
            status_rows.append({
                "Sold Party": party, "Generation_Status": "REVIEW", "Generation_Reason": "No eligible deal mode/recipe",
                "Interval_Status": p.get("Interval_Status"), "Behavior_Type": behavior.get("Behavior_Type"),
            })
            continue
        for r in results:
            if not r.line_status.empty:
                r.line_status["Chosen_Mode"] = "YES" if r.mode == chosen.mode else "NO"
                line_parts.append(r.line_status)
            if not r.candidates.empty:
                r.candidates["Chosen_Mode"] = "YES" if r.mode == chosen.mode else "NO"
                candidate_parts.append(r.candidates)
        pair_added_count = int(chosen.selected.get("Pair_Added_Flag", pd.Series("NO", index=chosen.selected.index)).astype(str).str.upper().eq("YES").sum()) if not chosen.selected.empty else 0
        normal_selected_count = len(chosen.selected) - pair_added_count
        status = "GENERATED" if chosen.accepted else "REVIEW - INSUFFICIENT STOCK"
        if unique_allocation and chosen.accepted and not chosen.selected.empty:
            global_used.update(chosen.selected["Serial_Key"].dropna().astype(str))
        status_rows.append({
            "Sold Party": party,
            "Sales_Person": p.get("Sales_Person"),
            "Interval_Status": p.get("Interval_Status"),
            "Behavior_Type": behavior.get("Behavior_Type"),
            "Primary_Trigger": behavior.get("Primary_Trigger"),
            "Chosen_Deal_Mode": chosen.mode,
            "Target_Stone_Count": chosen.target,
            "Selected_Stone_Count": len(chosen.selected),
            "Normal_Selected_Stone_Count": normal_selected_count,
            "Pair_Added_Stone_Count": pair_added_count,
            "Stock_Coverage": chosen.coverage,
            "Core_Stock_Coverage": chosen.core_coverage,
            "Generation_Status": status,
            "Generation_Reason": chosen.reason,
            "Global_Allocation_Mode": allocation_mode,
        })
        # Only accepted baskets belong in the final selected-stone output.
        # Partial attempts remain visible in candidate and recipe-fill review tables.
        if chosen.accepted and not chosen.selected.empty:
            selected = chosen.selected.copy()
            selected["Sales_Person"] = p.get("Sales_Person")
            selected["Interval_Status"] = p.get("Interval_Status")
            selected["Behavior_Type"] = behavior.get("Behavior_Type")
            selected_parts.append(selected)
    return (
        pd.DataFrame(status_rows),
        pd.concat(selected_parts, ignore_index=True) if selected_parts else pd.DataFrame(),
        pd.concat(line_parts, ignore_index=True) if line_parts else pd.DataFrame(),
        pd.concat(candidate_parts, ignore_index=True) if candidate_parts else pd.DataFrame(),
    )


def build_cross_party_conflicts(selected: pd.DataFrame) -> pd.DataFrame:
    if selected.empty or "Serial_Key" not in selected.columns:
        return pd.DataFrame(columns=["Serial_Key", "Serial No", "Party_Count", "Selection_Rows", "Parties", "Deal_Modes", "Selection_Sources", "Highest_Selection_Score"])
    work = selected[selected["Serial_Key"].astype(str).ne("")].copy()
    counts = work.groupby("Serial_Key")["Sold Party"].nunique()
    conflict_keys = set(counts[counts > 1].index)
    if not conflict_keys:
        return pd.DataFrame(columns=["Serial_Key", "Serial No", "Party_Count", "Selection_Rows", "Parties", "Deal_Modes", "Selection_Sources", "Highest_Selection_Score"])
    rows = []
    for serial, g in work[work["Serial_Key"].isin(conflict_keys)].groupby("Serial_Key", sort=False):
        rows.append({
            "Serial_Key": serial,
            "Serial No": first_nonblank(g.get("Serial No", pd.Series(dtype=object))),
            "Party_Count": g["Sold Party"].nunique(),
            "Selection_Rows": len(g),
            "Parties": " | ".join(sorted(set(g["Sold Party"].astype(str)))),
            "Deal_Modes": ",".join(sorted(set(g.get("Deal_Mode", pd.Series(dtype=str)).dropna().astype(str)))),
            "Selection_Sources": ",".join(sorted(set(g.get("Selection_Source", pd.Series(dtype=str)).dropna().astype(str)))),
            "Highest_Selection_Score": pd.to_numeric(g.get("Final_Selection_Score"), errors="coerce").max(),
        })
    return pd.DataFrame(rows).sort_values(["Party_Count", "Selection_Rows", "Highest_Selection_Score"], ascending=[False, False, False]).reset_index(drop=True)


# -----------------------------------------------------------------------------
# Excel output
# -----------------------------------------------------------------------------
def style_sheet(ws, header_row: int = 1) -> None:
    dark = "1F4E78"
    thin = Border(bottom=Side(style="thin", color="D9E2F3"))
    for cell in ws[header_row]:
        cell.fill = PatternFill("solid", fgColor=dark)
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = f"A{header_row + 1}"
    ws.sheet_view.showGridLines = False
    if ws.max_row >= header_row and ws.max_column:
        ws.auto_filter.ref = f"A{header_row}:{get_column_letter(ws.max_column)}{ws.max_row}"
    for row in ws.iter_rows(min_row=header_row + 1):
        for cell in row:
            cell.border = thin
            cell.alignment = Alignment(vertical="top", wrap_text=False)
    headers = [c.value for c in ws[header_row]]
    for idx in range(1, ws.max_column + 1):
        vals = [str(ws.cell(r, idx).value or "") for r in range(header_row, min(ws.max_row, 150) + 1)]
        width = min(max(max((len(x) for x in vals), default=8) + 2, 10), 45)
        ws.column_dimensions[get_column_letter(idx)].width = width
    for col_name in headers:
        if col_name and "Date" in str(col_name):
            j = headers.index(col_name) + 1
            for r in range(header_row + 1, ws.max_row + 1):
                ws.cell(r, j).number_format = "DD-MM-YYYY"
        if col_name and any(x in str(col_name) for x in ["Rate", "Share", "Coverage", "Progress"]):
            j = headers.index(col_name) + 1
            for r in range(header_row + 1, ws.max_row + 1):
                ws.cell(r, j).number_format = "0.0%"
    if "Generation_Status" in headers:
        j = headers.index("Generation_Status") + 1
        for r in range(header_row + 1, ws.max_row + 1):
            val = clean_text(ws.cell(r, j).value).upper()
            fill = "C6EFCE" if val == "GENERATED" else "FFF2CC" if "REVIEW" in val else "F2F2F2"
            ws.cell(r, j).fill = PatternFill("solid", fgColor=fill)
    if "Auto_Include" in headers:
        j = headers.index("Auto_Include") + 1
        letter = get_column_letter(j)
        ws.conditional_formatting.add(
            f"{letter}{header_row + 1}:{letter}{ws.max_row}",
            FormulaRule(formula=[f'{letter}{header_row + 1}="YES"'], fill=PatternFill("solid", fgColor="C6EFCE")),
        )


def write_dataframe(ws, df: pd.DataFrame) -> None:
    for c, name in enumerate(df.columns, 1):
        ws.cell(1, c, name)
    for r, row in enumerate(df.itertuples(index=False, name=None), 2):
        for c, value in enumerate(row, 1):
            ws.cell(r, c, safe_excel(value))
    style_sheet(ws)


def output_stock_columns(df: pd.DataFrame) -> list[str]:
    first = [c for c in OUTPUT_STOCK_COLUMNS if c in df.columns]
    return first + [c for c in df.columns if c not in first and not c.startswith("_")]


def create_analysis_workbook(path: Path, frames: dict[str, pd.DataFrame], run_summary: list[tuple[str, Any]]) -> None:
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("Dashboard")
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:H2")
    ws["A1"] = "DEALMAKER FULL TRIAL CONTROL"
    ws["A1"].fill = PatternFill("solid", fgColor="17365D")
    ws["A1"].font = Font(color="FFFFFF", bold=True, size=20)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A4"] = "Version"; ws["B4"] = VERSION
    ws["A5"] = "Purpose"; ws["B5"] = "Interval decides WHEN; party behavior and basket recipes decide WHAT; stock quotas decide WHICH stones."
    ws["A7"] = "Review order"
    ws["A8"] = "1. Party_Generation_Status"
    ws["A9"] = "2. Selected_Deal_Stones"
    ws["A10"] = "3. Basket_Recipe_Lines and Recipe_Fill_Status"
    ws["A11"] = "4. Current_Action_Candidates and Party_Action_Conversion"
    ws["A12"] = "5. Data_Quality"
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 100

    ws = wb.create_sheet("Run_Summary")
    write_dataframe(ws, pd.DataFrame(run_summary, columns=["Metric", "Value"]))

    order = [
        "Party_Generation_Status", "Selected_Deal_Stones", "Recipe_Fill_Status", "Party_Behavior_Profile",
        "Party_Deal_Summary", "Party_Deal_Mode_Summary", "Deal_Buckets", "Deal_Stone_Details",
        "Party_Action_Conversion", "Current_Action_Candidates", "Combination_Profile", "Basket_Recipe_Lines",
        "Cross_Party_Conflicts", "Stock_Sale_Overlap", "All_Stock_Candidates", "Excluded_Candidates", "Data_Quality",
    ]
    for name in order:
        df = frames.get(name, pd.DataFrame())
        ws = wb.create_sheet(name[:31])
        if name == "Selected_Deal_Stones" and not df.empty:
            df = df[output_stock_columns(df)]
        write_dataframe(ws, df)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def create_deal_workbooks(selected: pd.DataFrame, status: pd.DataFrame, output_dir: Path) -> tuple[Path | None, Path | None]:
    if selected.empty:
        return None, None
    combined_path = output_dir / "Generated_Deal_Baskets.xlsx"
    wb = Workbook(); wb.remove(wb.active)
    nav = wb.create_sheet("Navigation")
    nav.append(["Sold Party", "Sales Person", "Deal Mode", "Stone Count", "Status", "Go"])
    style_sheet(nav)
    nav_row = 2
    for party, g in selected.groupby("Sold Party"):
        title = re.sub(r"[\\/*?:\[\]]", "", str(party))[:31] or "Party"
        ws = wb.create_sheet(title)
        out = g[output_stock_columns(g)].copy()
        write_dataframe(ws, out)
        ws["A1"].hyperlink = "#'Navigation'!A1"
        srow = status[status["Sold Party"].eq(party)].iloc[0] if not status[status["Sold Party"].eq(party)].empty else pd.Series()
        nav.cell(nav_row, 1, party)
        nav.cell(nav_row, 2, srow.get("Sales_Person"))
        nav.cell(nav_row, 3, srow.get("Chosen_Deal_Mode"))
        nav.cell(nav_row, 4, len(g))
        nav.cell(nav_row, 5, srow.get("Generation_Status"))
        nav.cell(nav_row, 6, "Go")
        nav.cell(nav_row, 6).hyperlink = f"#'{title}'!A1"
        nav.cell(nav_row, 6).font = Font(color="0000FF", underline="single")
        nav_row += 1
    wb.save(combined_path)

    zip_path = output_dir / "Generated_Deal_Baskets_By_Salesperson.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for salesperson, sg in selected.groupby("Sales_Person", dropna=False):
            xwb = Workbook(); xwb.remove(xwb.active)
            xnav = xwb.create_sheet("Navigation")
            xnav.append(["Sold Party", "Deal Mode", "Stone Count", "Go"])
            for i, (party, g) in enumerate(sg.groupby("Sold Party"), 2):
                title = re.sub(r"[\\/*?:\[\]]", "", str(party))[:31] or "Party"
                ws = xwb.create_sheet(title)
                write_dataframe(ws, g[output_stock_columns(g)])
                ws["A1"].hyperlink = "#'Navigation'!A1"
                xnav.cell(i, 1, party); xnav.cell(i, 2, g["Deal_Mode"].iloc[0]); xnav.cell(i, 3, len(g)); xnav.cell(i, 4, "Go")
                xnav.cell(i, 4).hyperlink = f"#'{title}'!A1"
            style_sheet(xnav)
            buf = BytesIO(); xwb.save(buf)
            safe_name = re.sub(r"[\\/*?:\[\]<>|\"]", "", clean_text(salesperson, "Unknown Salesperson")) or "Unknown Salesperson"
            zf.writestr(f"{safe_name}.xlsx", buf.getvalue())
    return combined_path, zip_path


# -----------------------------------------------------------------------------
# Main run
# -----------------------------------------------------------------------------
def run(args: argparse.Namespace) -> None:
    sales_path = Path(args.sales)
    actions_path = Path(args.actions)
    stock_path = Path(args.stock)
    control_path = Path(args.control)
    output_dir = Path(args.output_dir)
    interval_engine_path = Path(args.interval_engine)
    output_dir.mkdir(parents=True, exist_ok=True)
    _t0 = time.perf_counter()
    faulthandler.enable()
    def _stage(label: str) -> None:
        print(f"[stage] {label}: {time.perf_counter() - _t0:.1f}s", flush=True)

    interval = load_module(interval_engine_path, "dealmaker_interval_engine_v6")
    settings, overrides, removed, exclusions = load_control(control_path)
    quality_log: list[dict[str, Any]] = []
    sales_mapping = load_manual_mapping(control_path, "SALES")
    actions_mapping = load_manual_mapping(control_path, "ACTIONS")
    stock_mapping = load_manual_mapping(control_path, "STOCK")
    party_aliases = load_party_aliases(control_path)

    sales_all = interval.prepare_dataset(
        sales_path, "Sales", quality_log, args.sales_sheet or None, args.sales_header_row, sales_mapping
    )
    actions = interval.prepare_dataset(
        actions_path, "Actions", quality_log, args.actions_sheet or None, args.actions_header_row, actions_mapping
    )
    _stage(f"inputs prepared (sales={len(sales_all)}, actions={len(actions)})")
    sales_all = apply_party_aliases(sales_all, party_aliases, "Sales", quality_log)
    actions = apply_party_aliases(actions, party_aliases, "Actions", quality_log)
    sales_all = interval.dedupe_business_rows(sales_all, "Sales", quality_log)
    actions = interval.dedupe_business_rows(actions, "Actions", quality_log)
    actions = canonicalize_action_types(actions, settings)
    as_of = interval.determine_as_of(settings, sales_all, actions)
    _stage(f"business duplicates removed (sales={len(sales_all)}, actions={len(actions)})")

    sale_types = parse_csv_set(settings.get("SALE_TYPE_VALUES", "SALE,SALES"))
    bid_types = parse_csv_set(settings.get("BID_TYPE_VALUES", "BID"))
    sales_only = sales_all[sales_all["Type"].isin(sale_types)].copy()
    bids = sales_all[sales_all["Type"].isin(bid_types)].copy()
    if sales_only.empty:
        raise ValueError(f"No completed sale rows found. SALE_TYPE_VALUES={sorted(sale_types)}")

    min_stones = int(float(settings.get("MIN_DEAL_STONES", 6)))
    max_days = int(float(settings.get("MAX_DEAL_DAYS", 2)))
    deals, _ = interval.build_deal_buckets(sales_only, min_stones, max_days)
    _stage(f"deal buckets built ({len(deals)})")
    party_summary, _ = interval.build_party_summary(sales_all, sales_only, bids, actions, deals, overrides, removed, settings, as_of)
    mode_summary = build_mode_summary(deals, party_summary, settings, as_of)
    conversion_window = int(float(settings.get("ACTION_CONVERSION_WINDOW_DAYS", 30)))
    action_conversion = build_action_conversion(actions, sales_only, deals, conversion_window)
    _stage(f"party summary, modes and action conversion built ({len(action_conversion)})")
    behavior_profiles = build_behavior_profiles(party_summary, mode_summary, action_conversion, settings)
    detail = add_deal_membership(sales_only, deals)
    detail = enrich_pattern_fields(detail) if not detail.empty else detail
    _stage(f"deal detail built (detail={len(detail)})")
    combination_profiles = build_combination_profiles(detail, deals, settings)
    recipes = build_recipe_lines(combination_profiles, mode_summary, settings)
    current_actions = build_current_actions(actions, deals, action_conversion, settings, as_of)
    _stage(f"behavior, combinations, recipes and current actions built (recipes={len(recipes)}, current={len(current_actions)})")

    stock, duplicate_stock, unavailable_stock = prepare_stock(
        stock_path, interval, settings, quality_log, args.stock_sheet or None, args.stock_header_row, stock_mapping
    )
    historical_sale_stock = stock[stock["Serial_Key"].isin(set(sales_only["Serial_Key"]))].copy()
    pair_sale_excluded = stock.iloc[0:0].copy()
    if yes(settings.get("EXCLUDE_COMPLETED_SALE_SERIALS_FROM_STOCK", "YES")) and not historical_sale_stock.empty:
        stock = stock[~stock["Serial_Key"].isin(set(historical_sale_stock["Serial_Key"]))].copy()
        stock, pair_sale_excluded = cascade_pair_exclusions(
            stock,
            historical_sale_stock,
            "PAIR REMOVED: RELATED MEMBER FOUND IN COMPLETED-SALE HISTORY",
        )
        if not pair_sale_excluded.empty:
            unavailable_stock = pd.concat([unavailable_stock, pair_sale_excluded], ignore_index=True, sort=False)
        stock = stock.reset_index(drop=True)
    quality_log.extend([
        {
            "Dataset": "Stock",
            "Check": "Completed-SALE history serials excluded from candidate stock",
            "Value": len(historical_sale_stock) if yes(settings.get("EXCLUDE_COMPLETED_SALE_SERIALS_FROM_STOCK", "YES")) else 0,
            "Severity": "WARNING" if len(historical_sale_stock) else "OK",
        },
        {
            "Dataset": "Stock",
            "Check": "Related pair members excluded with completed-sale serials",
            "Value": len(pair_sale_excluded),
            "Severity": "WARNING" if len(pair_sale_excluded) else "OK",
        },
    ])
    _stage(f"stock prepared ({len(stock)} eligible; {len(historical_sale_stock)} historical-sale overlaps)")
    status, selected, line_status, all_candidates = generate_deals(
        party_summary, behavior_profiles, mode_summary, recipes, current_actions, stock,
        overrides, removed, exclusions, settings,
    )
    excluded_candidates = all_candidates[all_candidates.get("Candidate_Status", pd.Series(index=all_candidates.index, dtype=str)).eq("EXCLUDED")].copy() if not all_candidates.empty else pd.DataFrame()
    cross_party_conflicts = build_cross_party_conflicts(selected)
    max_candidate_rows = int(float(settings.get("MAX_CANDIDATE_ROWS_IN_WORKBOOK", 50000)))
    workbook_candidates = all_candidates
    full_candidate_path = None
    if len(all_candidates) > max_candidate_rows:
        full_candidate_path = output_dir / "All_Stock_Candidates_Full.csv.gz"
        all_candidates.to_csv(full_candidate_path, index=False, compression="gzip")
        workbook_candidates = all_candidates.sort_values(
            ["Candidate_Status", "Final_Selection_Score"], ascending=[True, False]
        ).head(max_candidate_rows).copy()
    _stage(f"deal generation completed (selected={len(selected)}, candidates={len(all_candidates)})")

    data_quality = pd.DataFrame(quality_log)
    extra_quality = [
        {"Dataset": "Sales", "Check": "Completed SALE rows", "Value": len(sales_only), "Severity": "OK" if len(sales_only) else "ERROR"},
        {"Dataset": "Sales", "Check": "BID rows kept as intent only", "Value": len(bids), "Severity": "OK"},
        {"Dataset": "Deals", "Check": "Real deal buckets", "Value": len(deals), "Severity": "OK" if len(deals) else "WARNING"},
        {"Dataset": "Actions", "Check": "Action types found", "Value": ", ".join(sorted(actions["Type"].dropna().unique())), "Severity": "OK"},
        {"Dataset": "Actions", "Check": "Unmapped action types", "Value": ", ".join(sorted(set(actions["Type"].dropna()) - set(ACTION_TYPES))) or "None", "Severity": "WARNING" if (set(actions["Type"].dropna()) - set(ACTION_TYPES)) else "OK"},
        {"Dataset": "Actions", "Check": "HOLD available", "Value": int(actions["Type"].eq("HOLD").sum()), "Severity": "OK" if actions["Type"].eq("HOLD").any() else "WARNING"},
        {"Dataset": "Actions", "Check": "OFFER available", "Value": int(actions["Type"].eq("OFFER").sum()), "Severity": "OK" if actions["Type"].eq("OFFER").any() else "WARNING"},
        {"Dataset": "Stock", "Check": "Eligible stock rows after status and serial dedupe", "Value": len(stock), "Severity": "OK"},
        {"Dataset": "Stock", "Check": "Rows removed by status", "Value": len(unavailable_stock), "Severity": "WARNING" if len(unavailable_stock) else "OK"},
        {"Dataset": "Stock", "Check": "Duplicate serial rows removed", "Value": len(duplicate_stock), "Severity": "WARNING" if len(duplicate_stock) else "OK"},
        {"Dataset": "Stock", "Check": "Current-stock serials also present in completed-sale history", "Value": len(historical_sale_stock), "Severity": "WARNING" if len(historical_sale_stock) else "OK"},
        {"Dataset": "Output", "Check": "Cross-party selected-serial conflicts", "Value": len(cross_party_conflicts), "Severity": "WARNING" if len(cross_party_conflicts) else "OK"},
        {"Dataset": "Output", "Check": "Full candidate table exported separately", "Value": str(full_candidate_path.name) if full_candidate_path else "No", "Severity": "INFO"},
        {"Dataset": "Output", "Check": "Generated parties", "Value": int(status["Generation_Status"].eq("GENERATED").sum()) if not status.empty else 0, "Severity": "OK"},
    ]
    data_quality = pd.concat([data_quality, pd.DataFrame(extra_quality)], ignore_index=True)

    frames = {
        "Party_Generation_Status": status,
        "Selected_Deal_Stones": selected,
        "Recipe_Fill_Status": line_status,
        "Party_Behavior_Profile": behavior_profiles,
        "Party_Deal_Summary": party_summary,
        "Party_Deal_Mode_Summary": mode_summary,
        "Deal_Buckets": deals,
        "Deal_Stone_Details": detail,
        "Party_Action_Conversion": action_conversion,
        "Current_Action_Candidates": current_actions,
        "Combination_Profile": combination_profiles,
        "Basket_Recipe_Lines": recipes,
        "Cross_Party_Conflicts": cross_party_conflicts,
        "Stock_Sale_Overlap": historical_sale_stock,
        "All_Stock_Candidates": workbook_candidates,
        "Excluded_Candidates": excluded_candidates,
        "Data_Quality": data_quality,
    }
    summary = [
        ("Pipeline version", VERSION),
        ("Data as-of date", as_of),
        ("Sales input rows", len(sales_all)),
        ("Completed SALE rows", len(sales_only)),
        ("BID rows (intent only)", len(bids)),
        ("Action input rows", len(actions)),
        ("Stock input rows after serial dedupe", len(stock)),
        ("Real deal buckets", len(deals)),
        ("Deal-maker parties", int(party_summary["Deal_Maker_Class"].eq("DEAL MAKER").sum())),
        ("Interval-eligible parties", int(party_summary["Auto_Include"].eq("YES").sum())),
        ("Parties with generated baskets", int(status["Generation_Status"].eq("GENERATED").sum()) if not status.empty else 0),
        ("Selected stock stones", len(selected)),
        ("Historical-sale stock overlaps excluded", len(historical_sale_stock) if yes(settings.get("EXCLUDE_COMPLETED_SALE_SERIALS_FROM_STOCK", "YES")) else 0),
        ("Cross-party serial conflicts", len(cross_party_conflicts)),
        ("Real deal definition", f">= {min_stones} unique completed-SALE stones in 1-{max_days} continuous day(s)"),
        ("Production safety note", "Review Data_Quality and Party_Generation_Status before using any generated basket."),
    ]
    analysis_path = output_dir / args.analysis_name
    create_analysis_workbook(analysis_path, frames, summary)
    _stage("analysis workbook written")
    combined, zipped = create_deal_workbooks(selected, status, output_dir)
    _stage("deal workbooks written")

    print(f"Analysis workbook: {analysis_path}")
    if combined:
        print(f"Generated deal workbook: {combined}")
    if zipped:
        print(f"Salesperson ZIP: {zipped}")
    print("\nParty generation status:")
    if not status.empty:
        cols = [c for c in ["Sold Party", "Interval_Status", "Behavior_Type", "Chosen_Deal_Mode", "Target_Stone_Count", "Selected_Stone_Count", "Stock_Coverage", "Generation_Status"] if c in status.columns]
        print(status[cols].to_string(index=False))


def parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Full interval, behavior, recipe and stock deal generator.")
    p.add_argument("--sales", help="Full Sales/BID history Excel file.")
    p.add_argument("--actions", help="Full actions history Excel file.")
    p.add_argument("--stock", help="Current stock Excel file.")
    p.add_argument("--control", default="Dealmaker_Full_Control_v6.xlsx", help="Editable control workbook.")
    p.add_argument("--interval-engine", default="Dealmaker_Step1_Interval_Engine_v10.py", help="Interval engine script path.")
    p.add_argument("--sales-sheet", default="", help="Optional exact Sales worksheet name; blank auto-detects.")
    p.add_argument("--actions-sheet", default="", help="Optional exact Actions worksheet name; blank auto-detects.")
    p.add_argument("--stock-sheet", default="", help="Optional exact Stock worksheet name; blank auto-detects.")
    p.add_argument("--sales-header-row", type=int, default=None, help="Optional zero-based Sales header row.")
    p.add_argument("--actions-header-row", type=int, default=None, help="Optional zero-based Actions header row.")
    p.add_argument("--stock-header-row", type=int, default=None, help="Optional zero-based Stock header row.")
    p.add_argument("--analysis-name", default="Dealmaker_Production_Analysis_v6.xlsx", help="Analysis workbook filename.")
    p.add_argument("--output-dir", default="dealmaker_output", help="Output folder.")
    p.add_argument("--create-control-only", action="store_true", help="Create control workbook and exit.")
    return p


if __name__ == "__main__":
    args = parser().parse_args()
    if args.create_control_only:
        create_control_workbook(Path(args.control))
        print(f"Created control workbook: {args.control}")
    else:
        missing = [name for name in ["sales", "actions", "stock"] if not getattr(args, name)]
        if missing:
            raise SystemExit("Missing required arguments for a full run: " + ", ".join("--" + x for x in missing))
        run(args)
