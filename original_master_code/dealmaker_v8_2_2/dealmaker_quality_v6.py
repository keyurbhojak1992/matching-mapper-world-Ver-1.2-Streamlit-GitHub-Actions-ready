"""Configurable, fail-closed quality normalization for Dealmaker v6.

The workbook is the source of truth.  New nonblank values are never silently
converted to NONE: they are reported as UNMAPPED and block production when the
control setting requires a complete domain.  Blank values remain UNKNOWN.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd

DEFAULT_DOMAIN_ROWS = [
    ("Shade", "NONE", 0, "YES"),
    ("Shade", "INSPECTION ADVISABLE", 1, "YES"),
    ("Shade", "MIX TINGE 1", 2, "YES"),
    ("Shade", "MIX TINGE 2", 3, "YES"),
    ("Shade", "MIX TINGE 3", 4, "YES"),
    ("Shade", "LIGHT GREEN", 5, "NO"),
    ("Shade", "LIGHT BROWN", 6, "YES"),
    ("Shade", "BROWN", 7, "YES"),
    ("Shade", "STRONG BROWN", 8, "YES"),
    ("Shade", "VERY STRONG BROWN", 9, "YES"),
    ("Milky", "NONE", 0, "YES"),
    ("Milky", "LIGHT MILKY", 1, "YES"),
    ("Milky", "LIGHT CLOUD", 2, "YES"),
    ("Milky", "MILKY", 3, "YES"),
    ("Milky", "STRONG MILKY", 4, "YES"),
    ("Milky", "VERY STRONG MILKY", 5, "YES"),
    ("Milky", "INSPECTION ADVISABLE", 6, "YES"),
    ("Milky", "CLOUD", 7, "NO"),
    ("Milky", "MEDIUM CLOUD", 8, "NO"),
    ("Milky", "I A CLOUD", 9, "NO"),
    ("Open", "NONE", 0, "YES"),
    ("Open", "HAIR LINE", 1, "YES"),
    ("Open", "SMALL OPEN", 2, "YES"),
    ("Open", "MEDIUM OPEN", 3, "YES"),
    ("Open", "HEAVY OPEN", 4, "YES"),
    ("Blk", "NONE", 0, "YES"),
    ("Blk", "PIN POINT BLK", 1, "YES"),
    ("Blk", "MINOR BLK", 2, "YES"),
    ("Blk", "NOTICEABLE BLACK", 3, "YES"),
    ("Blk", "MEDIUM BLK", 4, "YES"),
    ("Blk", "HEAVY BLK", 5, "YES"),
]

DEFAULT_ALIAS_ROWS = [
    ("Shade", "NO SHADE", "NONE", "YES"),
    ("Shade", "NIL", "NONE", "YES"),
    ("Shade", "I A", "INSPECTION ADVISABLE", "YES"),
    ("Shade", "I.A.", "INSPECTION ADVISABLE", "YES"),
    ("Milky", "NO MILKY", "NONE", "YES"),
    ("Milky", "NIL", "NONE", "YES"),
    ("Milky", "I A", "INSPECTION ADVISABLE", "YES"),
    ("Milky", "I.A.", "INSPECTION ADVISABLE", "YES"),
    ("Open", "HVY OPEN", "HEAVY OPEN", "YES"),
    ("Open", "HEAVY OPENING", "HEAVY OPEN", "YES"),
    ("Open", "NO OPEN", "NONE", "YES"),
    ("Open", "NIL", "NONE", "YES"),
    ("Blk", "NOTICABLE BLACK", "NOTICEABLE BLACK", "YES"),
    ("Blk", "NOTICEABLE BLK", "NOTICEABLE BLACK", "YES"),
    ("Blk", "NOTICABLE BLK", "NOTICEABLE BLACK", "YES"),
    ("Blk", "PINPOINT BLK", "PIN POINT BLK", "YES"),
    ("Blk", "PIN POINT BLACK", "PIN POINT BLK", "YES"),
    ("Blk", "MINOR BLACK", "MINOR BLK", "YES"),
    ("Blk", "MEDIUM BLACK", "MEDIUM BLK", "YES"),
    ("Blk", "HEAVY BLACK", "HEAVY BLK", "YES"),
    ("Blk", "NO BLACK", "NONE", "YES"),
    ("Blk", "NIL", "NONE", "YES"),
]


def clean_quality_text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return " ".join(str(value).strip().upper().split())


def _yes(value: Any) -> bool:
    return clean_quality_text(value) in {"YES", "Y", "TRUE", "1", "ACTIVE"}


@dataclass
class QualityConfig:
    domain: pd.DataFrame
    aliases: pd.DataFrame
    orders: dict[str, list[str]]
    score_maps: dict[str, dict[str, int]]
    reverse_maps: dict[str, dict[int, str]]
    auto_maps: dict[str, dict[str, bool]]
    alias_maps: dict[str, dict[str, str]]
    validation_issues: list[dict[str, Any]]


def _default_domain() -> pd.DataFrame:
    return pd.DataFrame(DEFAULT_DOMAIN_ROWS, columns=["Dimension", "Canonical_Value", "Severity_Score", "Auto_Eligible"])


def _default_aliases() -> pd.DataFrame:
    return pd.DataFrame(DEFAULT_ALIAS_ROWS, columns=["Dimension", "Source_Value", "Canonical_Value", "Active"])


def _read_sheet(control: Path, names: list[str]) -> pd.DataFrame | None:
    try:
        xls = pd.ExcelFile(control)
    except Exception:
        return None
    for name in names:
        if name in xls.sheet_names:
            return pd.read_excel(control, sheet_name=name)
    return None


def load_quality_config(control: str | Path | None = None) -> QualityConfig:
    issues: list[dict[str, Any]] = []
    domain = _default_domain()
    aliases = _default_aliases()
    if control:
        control_path = Path(control)
        raw_domain = _read_sheet(control_path, ["Quality_Domain", "Quality_Scales"])
        if raw_domain is not None and not raw_domain.empty:
            rename = {"Score": "Severity_Score", "Automatic_Use": "Auto_Eligible"}
            raw_domain = raw_domain.rename(columns=rename)
            needed = {"Dimension", "Canonical_Value", "Severity_Score", "Auto_Eligible"}
            if needed.issubset(raw_domain.columns):
                domain = raw_domain[list(needed)].copy()
            else:
                issues.append({"Severity": "ERROR", "Dataset": "Control", "Issue": "Quality domain columns missing", "Value": sorted(needed - set(raw_domain.columns))})
        raw_alias = _read_sheet(control_path, ["Quality_Value_Aliases"])
        if raw_alias is not None and not raw_alias.empty:
            needed = {"Dimension", "Source_Value", "Canonical_Value"}
            if needed.issubset(raw_alias.columns):
                if "Active" not in raw_alias.columns:
                    raw_alias["Active"] = "YES"
                aliases = raw_alias[["Dimension", "Source_Value", "Canonical_Value", "Active"]].copy()
            else:
                issues.append({"Severity": "ERROR", "Dataset": "Control", "Issue": "Quality alias columns missing", "Value": sorted(needed - set(raw_alias.columns))})

    domain["Dimension"] = domain["Dimension"].map(clean_quality_text).str.title().replace({"Blk": "Blk", "Open": "Open", "Shade": "Shade", "Milky": "Milky"})
    domain["Canonical_Value"] = domain["Canonical_Value"].map(clean_quality_text)
    domain["Severity_Score"] = pd.to_numeric(domain["Severity_Score"], errors="coerce")
    domain["Auto_Eligible"] = domain["Auto_Eligible"].map(lambda x: "YES" if _yes(x) else "NO")
    domain = domain[domain["Dimension"].isin(["Shade", "Milky", "Open", "Blk"]) & domain["Canonical_Value"].ne("")].copy()
    domain = domain.drop_duplicates(["Dimension", "Canonical_Value"], keep="last").sort_values(["Dimension", "Severity_Score", "Canonical_Value"]).reset_index(drop=True)

    for dim in ["Shade", "Milky", "Open", "Blk"]:
        part = domain[domain["Dimension"].eq(dim)]
        if part.empty:
            issues.append({"Severity": "ERROR", "Dataset": "Control", "Issue": "Missing quality dimension", "Dimension": dim})
        if part["Severity_Score"].isna().any():
            issues.append({"Severity": "ERROR", "Dataset": "Control", "Issue": "Blank/non-numeric quality score", "Dimension": dim})
        if "NONE" not in set(part["Canonical_Value"]):
            issues.append({"Severity": "ERROR", "Dataset": "Control", "Issue": "Quality dimension lacks NONE", "Dimension": dim})

    aliases["Dimension"] = aliases["Dimension"].map(clean_quality_text).str.title().replace({"Blk": "Blk", "Open": "Open", "Shade": "Shade", "Milky": "Milky"})
    aliases["Source_Value"] = aliases["Source_Value"].map(clean_quality_text)
    aliases["Canonical_Value"] = aliases["Canonical_Value"].map(clean_quality_text)
    aliases["Active"] = aliases["Active"].map(lambda x: "YES" if _yes(x) else "NO")
    aliases = aliases[aliases["Active"].eq("YES") & aliases["Source_Value"].ne("")].drop_duplicates(["Dimension", "Source_Value"], keep="last").reset_index(drop=True)

    domain_values = {(r.Dimension, r.Canonical_Value) for r in domain.itertuples(index=False)}
    for row in aliases.itertuples(index=False):
        if (row.Dimension, row.Canonical_Value) not in domain_values:
            issues.append({"Severity": "ERROR", "Dataset": "Control", "Issue": "Alias target is not canonical", "Dimension": row.Dimension, "Value": row.Source_Value, "Target": row.Canonical_Value})

    orders: dict[str, list[str]] = {}
    score_maps: dict[str, dict[str, int]] = {}
    reverse_maps: dict[str, dict[int, str]] = {}
    auto_maps: dict[str, dict[str, bool]] = {}
    alias_maps: dict[str, dict[str, str]] = {}
    for dim in ["Shade", "Milky", "Open", "Blk"]:
        part = domain[domain["Dimension"].eq(dim)].sort_values(["Severity_Score", "Canonical_Value"])
        orders[dim] = part["Canonical_Value"].tolist()
        score_maps[dim] = {r.Canonical_Value: int(r.Severity_Score) for r in part.itertuples(index=False)}
        reverse_maps[dim] = {}
        for r in part.itertuples(index=False):
            reverse_maps[dim].setdefault(int(r.Severity_Score), r.Canonical_Value)
        auto_maps[dim] = {r.Canonical_Value: _yes(r.Auto_Eligible) for r in part.itertuples(index=False)}
        alias_maps[dim] = {r.Source_Value: r.Canonical_Value for r in aliases[aliases["Dimension"].eq(dim)].itertuples(index=False)}
    return QualityConfig(domain, aliases, orders, score_maps, reverse_maps, auto_maps, alias_maps, issues)


_ACTIVE_CONFIG = load_quality_config(None)
QUALITY_ORDERS = _ACTIVE_CONFIG.orders
QUALITY_MAPS = _ACTIVE_CONFIG.score_maps
QUALITY_REVERSE = _ACTIVE_CONFIG.reverse_maps
QUALITY_VALUE_ALIASES = _ACTIVE_CONFIG.alias_maps


def configure_quality(control: str | Path | None) -> QualityConfig:
    global _ACTIVE_CONFIG, QUALITY_ORDERS, QUALITY_MAPS, QUALITY_REVERSE, QUALITY_VALUE_ALIASES
    _ACTIVE_CONFIG = load_quality_config(control)
    QUALITY_ORDERS = _ACTIVE_CONFIG.orders
    QUALITY_MAPS = _ACTIVE_CONFIG.score_maps
    QUALITY_REVERSE = _ACTIVE_CONFIG.reverse_maps
    QUALITY_VALUE_ALIASES = _ACTIVE_CONFIG.alias_maps
    return _ACTIVE_CONFIG


def canonical_quality(value: Any, dimension: str, blank_as_none: bool = False) -> str:
    raw = clean_quality_text(value)
    if not raw:
        return "NONE" if blank_as_none else ""
    return _ACTIVE_CONFIG.alias_maps.get(dimension, {}).get(raw, raw)


def quality_score(value: Any, dimension: str, blank_as_none: bool = False) -> float:
    canonical = canonical_quality(value, dimension, blank_as_none=blank_as_none)
    score = _ACTIVE_CONFIG.score_maps.get(dimension, {}).get(canonical)
    return float(score) if score is not None else float("nan")


def quality_text(score: Any, dimension: str) -> str:
    if score is None or pd.isna(score):
        return "UNKNOWN"
    return _ACTIVE_CONFIG.reverse_maps.get(dimension, {}).get(int(round(float(score))), "UNKNOWN")


def quality_auto_eligible(value: Any, dimension: str) -> bool:
    canonical = canonical_quality(value, dimension, blank_as_none=False)
    return bool(canonical and _ACTIVE_CONFIG.auto_maps.get(dimension, {}).get(canonical, False))


def domain_status(value: Any, dimension: str) -> tuple[str, str, str]:
    raw = clean_quality_text(value)
    if not raw:
        return "", "UNKNOWN", "UNKNOWN"
    canonical = _ACTIVE_CONFIG.alias_maps.get(dimension, {}).get(raw, raw)
    if canonical not in _ACTIVE_CONFIG.score_maps.get(dimension, {}):
        return raw, canonical, "UNMAPPED"
    if not _ACTIVE_CONFIG.auto_maps.get(dimension, {}).get(canonical, False):
        return raw, canonical, "REVIEW_ONLY"
    return raw, canonical, "ALIASED" if raw != canonical else "CANONICAL"


def domain_audit(frame: pd.DataFrame, dataset: str, config: QualityConfig | None = None) -> pd.DataFrame:
    config = config or _ACTIVE_CONFIG
    columns = {
        "Shade": ["Shade"],
        "Milky": ["Milky"],
        "Open": ["TOIN_NAME", "SOIN_NAME", "GIRDLE_OPEN_NAME", "PAV_OPEN_NAME"],
        "Blk": ["TBINC_NAME", "SBINC_NAME"],
    }
    rows: list[dict[str, Any]] = []
    for dim, fields in columns.items():
        for field in fields:
            if field not in frame.columns:
                continue
            counts = frame[field].map(clean_quality_text).value_counts(dropna=False)
            for raw, count in counts.items():
                source = "" if raw is None else str(raw)
                if not source:
                    canonical, status = "UNKNOWN", "UNKNOWN"
                else:
                    canonical = config.alias_maps.get(dim, {}).get(source, source)
                    if canonical not in config.score_maps.get(dim, {}):
                        status = "UNMAPPED"
                    elif not config.auto_maps.get(dim, {}).get(canonical, False):
                        status = "REVIEW_ONLY"
                    elif source != canonical:
                        status = "ALIASED"
                    else:
                        status = "CANONICAL"
                rows.append({"Dataset": dataset, "Dimension": dim, "Column": field, "Raw_Value": source or "<BLANK>", "Canonical_Value": canonical, "Count": int(count), "Status": status})
    return pd.DataFrame(rows)
