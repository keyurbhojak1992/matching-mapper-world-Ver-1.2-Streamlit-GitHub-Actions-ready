#!/usr/bin/env python3
"""Robust input discovery and column mapping for Dealmaker production files."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
from typing import Any

import pandas as pd


@dataclass
class InputMetadata:
    dataset: str
    path: str
    file_type: str
    sheet_name: str
    header_row_zero_based: int
    row_count: int
    column_count: int
    raw_columns: list[str]
    standardized_columns: list[str]
    resolved_mapping: dict[str, str]
    missing_required: list[str]


def clean_name(value: Any) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\n", " ").replace("\r", " ")
    return re.sub(r"\s+", " ", text).strip()


def header_key(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", clean_name(value).lower())


def _alias_lookup(aliases: dict[str, list[str]], manual_mapping: dict[str, str] | None = None) -> dict[str, str]:
    lookup: dict[str, str] = {}
    for standard, values in aliases.items():
        lookup[header_key(standard)] = standard
        for value in values:
            lookup[header_key(value)] = standard
    for source, standard in (manual_mapping or {}).items():
        if clean_name(source) and clean_name(standard):
            lookup[header_key(source)] = clean_name(standard)
    return lookup


def standardize_columns(
    df: pd.DataFrame,
    aliases: dict[str, list[str]],
    manual_mapping: dict[str, str] | None = None,
) -> tuple[pd.DataFrame, dict[str, str]]:
    lookup = _alias_lookup(aliases, manual_mapping)
    raw_columns = [clean_name(c) for c in df.columns]
    mapped_columns = [lookup.get(header_key(c), clean_name(c)) for c in raw_columns]
    mapping = {raw: mapped for raw, mapped in zip(raw_columns, mapped_columns)}
    out = df.copy()
    out.columns = mapped_columns

    # Coalesce duplicate standardized columns. Real exports sometimes contain both
    # an old and a new label for the same field.
    if out.columns.duplicated().any():
        columns: list[str] = []
        data: dict[str, pd.Series] = {}
        for col in out.columns:
            if col in data:
                continue
            matches = out.loc[:, out.columns == col]
            if matches.shape[1] == 1:
                data[col] = matches.iloc[:, 0]
            else:
                merged = matches.iloc[:, 0]
                for idx in range(1, matches.shape[1]):
                    merged = merged.where(merged.notna() & merged.astype(str).str.strip().ne(""), matches.iloc[:, idx])
                data[col] = merged
            columns.append(col)
        out = pd.DataFrame({col: data[col] for col in columns})
    return out, mapping


def _read_delimited(path: Path, header_row: int) -> pd.DataFrame:
    sep = "\t" if path.suffix.lower() == ".tsv" else ","
    last_error: Exception | None = None
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
        try:
            return pd.read_csv(path, sep=sep, header=header_row, encoding=encoding, low_memory=False, keep_default_na=False, na_values=[])
        except Exception as exc:  # pragma: no cover - only reached for unusual encodings
            last_error = exc
    raise RuntimeError(f"Could not read delimited file {path}: {last_error}")


def read_dataset_table(
    path: str | Path,
    aliases: dict[str, list[str]],
    required: list[str] | set[str],
    dataset_name: str,
    sheet_name: str | None = None,
    header_row: int | None = None,
    manual_mapping: dict[str, str] | None = None,
    scan_header_rows: int = 15,
) -> tuple[pd.DataFrame, InputMetadata]:
    """Read the best table from Excel/CSV and standardize columns.

    ``header_row`` is zero-based, matching pandas. If omitted, Excel worksheets
    and the first 15 possible header rows are scored against required aliases.
    """
    source = Path(path)
    if not source.exists():
        raise FileNotFoundError(source)
    required_set = {clean_name(x) for x in required}
    suffix = source.suffix.lower()

    if suffix in {".csv", ".tsv"}:
        chosen_header = 0 if header_row is None else int(header_row)
        raw = _read_delimited(source, chosen_header)
        standardized, mapping = standardize_columns(raw, aliases, manual_mapping)
        missing = sorted(required_set - set(standardized.columns))
        meta = InputMetadata(
            dataset=dataset_name,
            path=str(source),
            file_type=suffix.lstrip(".").upper(),
            sheet_name="",
            header_row_zero_based=chosen_header,
            row_count=len(standardized),
            column_count=len(standardized.columns),
            raw_columns=[clean_name(c) for c in raw.columns],
            standardized_columns=[clean_name(c) for c in standardized.columns],
            resolved_mapping=mapping,
            missing_required=missing,
        )
        if missing:
            raise KeyError(f"{dataset_name} is missing required columns after mapping: {missing}")
        return standardized, meta

    if suffix not in {".xlsx", ".xlsm", ".xls"}:
        raise ValueError(f"Unsupported input format for {dataset_name}: {source.suffix}")

    xls = pd.ExcelFile(source)
    candidate_sheets = [sheet_name] if sheet_name else list(xls.sheet_names)
    for candidate in candidate_sheets:
        if candidate not in xls.sheet_names:
            raise KeyError(f"Worksheet '{candidate}' not found in {source.name}. Available: {xls.sheet_names}")

    best: tuple[tuple[int, int, int, int], str, int, pd.DataFrame, pd.DataFrame, dict[str, str]] | None = None
    headers = [int(header_row)] if header_row is not None else list(range(scan_header_rows))
    alias_lookup = _alias_lookup(aliases, manual_mapping)
    optional_standards = set(aliases)

    for sheet_index, candidate_sheet in enumerate(candidate_sheets):
        # Read a small raw sample once per worksheet. Reopening a 100k-row XLSX
        # for every possible header row is unnecessarily slow and can dominate
        # the full production preflight.
        try:
            sample = pd.read_excel(
                xls, sheet_name=candidate_sheet, header=None,
                nrows=scan_header_rows + 25, keep_default_na=False, na_values=[]
            )
        except Exception:
            continue
        for candidate_header in headers:
            if candidate_header >= len(sample):
                continue
            header_values = [clean_name(v) for v in sample.iloc[candidate_header].tolist()]
            preview = sample.iloc[candidate_header + 1:candidate_header + 26].copy()
            preview.columns = header_values
            if preview.empty and len(preview.columns) == 0:
                continue
            standardized_preview, mapping = standardize_columns(preview, aliases, manual_mapping)
            columns = set(standardized_preview.columns)
            required_hits = len(required_set & columns)
            optional_hits = len(optional_standards & columns)
            meaningful_headers = sum(1 for c in preview.columns if clean_name(c) and not clean_name(c).lower().startswith("unnamed"))
            # Higher required/optional hits, more meaningful headers, earlier header row/sheet.
            score = (required_hits, optional_hits, meaningful_headers, -candidate_header - sheet_index)
            if best is None or score > best[0]:
                best = (score, candidate_sheet, candidate_header, preview, standardized_preview, mapping)

    if best is None:
        raise RuntimeError(f"No readable worksheet/table found in {source.name}")

    _, chosen_sheet, chosen_header, _, _, _ = best
    raw = pd.read_excel(xls, sheet_name=chosen_sheet, header=chosen_header, keep_default_na=False, na_values=[])
    standardized, mapping = standardize_columns(raw, aliases, manual_mapping)
    standardized = standardized.dropna(how="all").reset_index(drop=True)
    missing = sorted(required_set - set(standardized.columns))
    meta = InputMetadata(
        dataset=dataset_name,
        path=str(source),
        file_type=suffix.lstrip(".").upper(),
        sheet_name=chosen_sheet,
        header_row_zero_based=chosen_header,
        row_count=len(standardized),
        column_count=len(standardized.columns),
        raw_columns=[clean_name(c) for c in raw.columns],
        standardized_columns=[clean_name(c) for c in standardized.columns],
        resolved_mapping=mapping,
        missing_required=missing,
    )
    if missing:
        recognized = sorted(c for c in standardized.columns if c in optional_standards)
        raise KeyError(
            f"{dataset_name} is missing required columns after automatic detection: {missing}. "
            f"Detected worksheet='{chosen_sheet}', header row={chosen_header + 1}, recognized={recognized}."
        )
    return standardized, meta


def _read_control_sheet(control_path: Path, sheet_name: str) -> pd.DataFrame:
    if not control_path.exists():
        return pd.DataFrame()
    try:
        xls = pd.ExcelFile(control_path)
        if sheet_name not in xls.sheet_names:
            return pd.DataFrame()
        first = pd.read_excel(control_path, sheet_name=sheet_name)
        if not first.empty and not all(str(c).lower().startswith("unnamed") for c in first.columns):
            return first
        return pd.read_excel(control_path, sheet_name=sheet_name, header=2)
    except Exception:
        return pd.DataFrame()


def load_manual_mapping(control_path: str | Path, dataset: str) -> dict[str, str]:
    """Return {source_column: standard_column} from optional Column_Mapping sheet."""
    df = _read_control_sheet(Path(control_path), "Column_Mapping")
    required = {"Dataset", "Standard_Column", "Source_Column"}
    if df.empty or not required.issubset(df.columns):
        return {}
    active = df.get("Active", pd.Series("YES", index=df.index)).fillna("YES").astype(str).str.strip().str.upper()
    wanted = df[active.isin({"YES", "Y", "TRUE", "1"})].copy()
    wanted = wanted[wanted["Dataset"].fillna("").astype(str).str.strip().str.upper().eq(dataset.upper())]
    mapping: dict[str, str] = {}
    for _, row in wanted.iterrows():
        source = clean_name(row.get("Source_Column"))
        standard = clean_name(row.get("Standard_Column"))
        if source and standard:
            mapping[source] = standard
    return mapping


def serial_key_series(series: pd.Series) -> pd.Series:
    """Normalize serials without deleting meaningful punctuation."""
    out = series.fillna("").astype(str).str.strip().str.upper()
    out = out.str.replace(r"\.0+$", "", regex=True)
    out = out.str.replace(r"\s+", "", regex=True)
    return out


def metadata_rows(meta: InputMetadata) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for raw, standardized in meta.resolved_mapping.items():
        rows.append({
            "Dataset": meta.dataset,
            "Source_File": Path(meta.path).name,
            "Worksheet": meta.sheet_name,
            "Header_Row_Excel": meta.header_row_zero_based + 1,
            "Source_Column": raw,
            "Standard_Column": standardized,
            "Mapped": "YES" if raw != standardized else "DIRECT/UNCHANGED",
        })
    return rows
