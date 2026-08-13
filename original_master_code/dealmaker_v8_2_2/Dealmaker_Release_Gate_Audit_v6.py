from __future__ import annotations

import argparse
import gzip
import io
import math
import re
import zipfile
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

VERSION = "8.2.2"


def clean(value: Any) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    return str(value).strip()


def yes(value: Any) -> bool:
    return clean(value).upper() in {"YES", "Y", "TRUE", "1"}


def load_settings(control: Path) -> dict[str, Any]:
    df = pd.read_excel(control, sheet_name="Settings")
    if df.empty:
        return {}
    key_col = df.columns[0]
    val_col = df.columns[1]
    return {
        clean(k): v
        for k, v in zip(df[key_col], df[val_col])
        if clean(k)
    }


def add_check(rows: list[dict[str, Any]], gate: str, passed: bool, value: Any, detail: str, severity: str = "ERROR") -> None:
    rows.append({
        "Gate": gate,
        "Status": "PASS" if passed else ("WARNING" if severity == "WARNING" else "FAIL"),
        "Value": value,
        "Severity": "OK" if passed else severity,
        "Detail": detail,
    })


def read_sheet(book: Path, sheet: str) -> pd.DataFrame:
    try:
        return pd.read_excel(book, sheet_name=sheet)
    except Exception:
        return pd.DataFrame()


def read_csv_safe(path: Path, **kwargs) -> pd.DataFrame:
    try:
        return pd.read_csv(path, **kwargs)
    except (pd.errors.EmptyDataError, pd.errors.ParserError, FileNotFoundError, OSError):
        return pd.DataFrame()


def ensure_columns(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    out = df.copy() if isinstance(df, pd.DataFrame) else pd.DataFrame()
    for column in columns:
        if column not in out.columns:
            out[column] = pd.Series(index=out.index, dtype="object")
    return out


def serial_rows_from_workbook_bytes(data: bytes) -> pd.DataFrame:
    # Salesperson-facing party sheets deliberately contain only the familiar
    # regular-stock columns. Reconciliation therefore uses Serial No and the
    # full party-name mapping stored on Navigation, rather than requiring
    # Dealmaker audit columns on every visible sheet.
    wb = load_workbook(io.BytesIO(data), read_only=False, data_only=False)
    rows: list[dict[str, str]] = []

    sheet_to_party: dict[str, str] = {}
    if "Navigation" in wb.sheetnames:
        nav = wb["Navigation"]
        for row_number in range(2, nav.max_row + 1):
            full_party = clean(nav.cell(row_number, 2).value or nav.cell(row_number, 1).value)
            link = nav.cell(row_number, 6).hyperlink
            target = clean(link.target if link else "")
            match = re.match(r"#?'([^']+)'!", target)
            if match and full_party:
                sheet_to_party[match.group(1)] = full_party

    for ws in wb.worksheets:
        if ws.title == "Navigation" or ws.title.startswith("Dealmaker") or ws.max_row < 2:
            continue
        headers = [clean(cell.value) for cell in ws[1]]
        index = {header: i for i, header in enumerate(headers)}
        serial_header = "Serial_Key" if "Serial_Key" in index else "Serial No" if "Serial No" in index else ""
        if not serial_header:
            continue
        for values in ws.iter_rows(min_row=2, values_only=True):
            serial = clean(values[index[serial_header]] if index[serial_header] < len(values) else "")
            if serial.endswith(".0"):
                serial = serial[:-2]
            if not serial:
                continue
            party = clean(values[index["Sold Party"]] if "Sold Party" in index and index["Sold Party"] < len(values) else sheet_to_party.get(ws.title, ws.title))
            rows.append({"Sold Party": party, "Serial_Key": serial, "Sheet": ws.title})
    return pd.DataFrame(rows)


def check_formula_errors_workbook(path_or_bytes: Path | bytes) -> int:
    source = io.BytesIO(path_or_bytes) if isinstance(path_or_bytes, bytes) else path_or_bytes
    wb = load_workbook(source, read_only=True, data_only=False)
    errors = {"#REF!", "#DIV/0!", "#VALUE!", "#N/A", "#NAME?", "#NUM!", "#NULL!"}
    count = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            for value in row:
                if isinstance(value, str) and value.strip().upper() in errors:
                    count += 1
    return count


def native_link_failures_from_workbook_bytes(data: bytes, workbook_name: str) -> pd.DataFrame:
    """Validate readable native stock hyperlinks, including long URL targets."""
    wb = load_workbook(io.BytesIO(data), read_only=False, data_only=False)
    link_headers = {"DETAIL", "IMAGE", "VIDEO", "CERTIFICATE", "PHYGITAL", "LAYOUT"}
    failures: list[dict[str, Any]] = []
    for ws in wb.worksheets:
        if ws.title == "Navigation" or ws.title.startswith("Dealmaker") or ws.max_row < 2:
            continue
        headers = {clean(cell.value): cell.column for cell in ws[1] if clean(cell.value)}
        for header in link_headers.intersection(headers):
            column = headers[header]
            for row_number in range(2, ws.max_row + 1):
                cell = ws.cell(row_number, column)
                value = clean(cell.value)
                formula_link = value.upper().startswith("=HYPERLINK(")
                raw_url = value.lower().startswith(("http://", "https://"))
                native_link = cell.hyperlink is not None
                # Unsupported non-URL source text may remain plain text. Only
                # URL-like or deliberately labeled cells are required to link.
                if not (formula_link or raw_url or native_link or value == header):
                    continue
                target = clean((cell.hyperlink.target or cell.hyperlink.location) if cell.hyperlink else "")
                if formula_link:
                    reason = "FORMULA_LINK_REMAINS"
                elif raw_url:
                    reason = "RAW_URL_VISIBLE"
                elif value != header:
                    reason = "VISIBLE_LABEL_MISMATCH"
                elif not target.lower().startswith(("http://", "https://")):
                    reason = "MISSING_OR_INVALID_NATIVE_TARGET"
                else:
                    continue
                failures.append({
                    "Workbook": workbook_name,
                    "Sheet": ws.title,
                    "Cell": cell.coordinate,
                    "Header": header,
                    "Value": value,
                    "Target_Length": len(target),
                    "Reason": reason,
                })
    return pd.DataFrame(failures)


HYPERLINK_FORMULA_RE = re.compile(
    r'^=HYPERLINK\(\s*"((?:[^"]|"")*)"\s*[,;]\s*"((?:[^"]|"")*)"\s*\)$',
    flags=re.IGNORECASE,
)
LINK_HEADERS = {"DETAIL", "IMAGE", "VIDEO", "CERTIFICATE", "PHYGITAL", "LAYOUT"}


def link_target(value: Any) -> str:
    text = clean(value)
    if text.lower().startswith(("http://", "https://")):
        return text
    match = HYPERLINK_FORMULA_RE.match(text)
    if not match:
        return ""
    target = match.group(1).replace('""', '"')
    return target if target.lower().startswith(("http://", "https://")) else ""


def expected_native_link_count(selected: pd.DataFrame) -> int:
    total = 0
    for header in LINK_HEADERS.intersection(selected.columns):
        total += int(selected[header].map(lambda value: bool(link_target(value))).sum())
    return total


def native_link_metrics_from_workbook_bytes(data: bytes, workbook_name: str) -> tuple[pd.DataFrame, dict[str, Any]]:
    failures = native_link_failures_from_workbook_bytes(data, workbook_name)
    wb = load_workbook(io.BytesIO(data), read_only=False, data_only=False)
    metrics = {
        "Workbook": workbook_name,
        "Native_HTTP_Links": 0,
        "Formula_Hyperlinks": 0,
        "Visible_Raw_URLs": 0,
        "Invalid_Native_Targets": 0,
        "Worksheet_XML_Hyperlink_Formulas": 0,
    }
    for ws in wb.worksheets:
        if ws.title == "Navigation" or ws.title.startswith("Dealmaker") or ws.max_row < 2:
            continue
        headers = {clean(cell.value): cell.column for cell in ws[1] if clean(cell.value)}
        for header in LINK_HEADERS.intersection(headers):
            column = headers[header]
            for row_number in range(2, ws.max_row + 1):
                cell = ws.cell(row_number, column)
                value = clean(cell.value)
                if value.upper().startswith("=HYPERLINK("):
                    metrics["Formula_Hyperlinks"] += 1
                if value.lower().startswith(("http://", "https://")):
                    metrics["Visible_Raw_URLs"] += 1
                if cell.hyperlink is None:
                    continue
                target = clean(cell.hyperlink.target or cell.hyperlink.location)
                if target.lower().startswith(("http://", "https://")):
                    metrics["Native_HTTP_Links"] += 1
                else:
                    metrics["Invalid_Native_Targets"] += 1
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        for name in archive.namelist():
            if name.startswith("xl/worksheets/") and name.endswith(".xml"):
                metrics["Worksheet_XML_Hyperlink_Formulas"] += archive.read(name).upper().count(b"HYPERLINK(")
    return failures, metrics


def style_workbook(path: Path) -> None:
    wb = load_workbook(path)
    header_fill = PatternFill("solid", fgColor="1F4E78")
    pass_fill = PatternFill("solid", fgColor="C6EFCE")
    fail_fill = PatternFill("solid", fgColor="F4CCCC")
    warn_fill = PatternFill("solid", fgColor="FFF2CC")
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.sheet_view.showGridLines = False
        if ws.max_row:
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = Font(color="FFFFFF", bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        headers = {clean(c.value): c.column for c in ws[1]} if ws.max_row else {}
        status_col = headers.get("Status")
        if status_col:
            for r in range(2, ws.max_row + 1):
                status = clean(ws.cell(r, status_col).value).upper()
                fill = pass_fill if status == "PASS" else warn_fill if status == "WARNING" else fail_fill
                for c in range(1, ws.max_column + 1):
                    ws.cell(r, c).fill = fill
        for c in range(1, ws.max_column + 1):
            max_len = max((len(clean(ws.cell(r, c).value)) for r in range(1, min(ws.max_row, 300) + 1)), default=0)
            ws.column_dimensions[get_column_letter(c)].width = min(max(max_len + 2, 12), 55)
    wb.save(path)


def run(output: Path, control: Path) -> tuple[pd.DataFrame, dict[str, pd.DataFrame]]:
    review_book = output / "Dealmaker_Production_Review_v6.xlsx"
    selected = read_csv_safe(output / "Selected_Deal_Stones_v6.csv", low_memory=False)
    status = read_csv_safe(output / "Party_Generation_Status_v6.csv", low_memory=False)
    line = read_csv_safe(output / "Recipe_Fill_Status_v6.csv", low_memory=False)
    profile = read_csv_safe(output / "Profile_Deal_Stone_Details_v6.csv.gz", low_memory=False)
    combos = read_csv_safe(output / "Combination_Profile_v6.csv.gz", low_memory=False)
    repeated_audit = read_csv_safe(output / "Repeated_Deal_Serial_Audit_v6.csv", low_memory=False)
    deals = read_sheet(review_book, "Deal_Buckets")
    data_quality = read_sheet(review_book, "Data_Quality")
    mode_summary = read_sheet(review_book, "Party_Deal_Mode_Summary")
    recipes = read_sheet(review_book, "Basket_Recipe_Lines")
    settings = load_settings(control)

    checks: list[dict[str, Any]] = []
    details: dict[str, pd.DataFrame] = {}

    # Normalize schemas before any gate logic. A valid run may contain zero
    # accepted baskets, but all gates must still complete and report cleanly.
    selected = ensure_columns(selected, [
        "Sold Party", "Deal_Mode", "Recipe_Line_ID", "Serial_Key",
        "Pair_Key", "Pair_Group_Members",
    ])
    status = ensure_columns(status, ["Sold Party", "Generation_Status", "Release_Status"])
    line = ensure_columns(line, [
        "Sold Party", "Deal_Mode", "Recipe_Line_ID",
        "Final_Globally_Allocated_Quantity", "Recipe_Fill_Reconciliation_Status",
    ])
    profile = ensure_columns(profile, [
        "Sold Party", "Deal_ID", "Serial_Key", "Deal_Mode",
        "Preferred_Combination_Key",
    ])
    combos = ensure_columns(combos, [
        "Sold Party", "Deal_Mode", "Preferred_Combination_Key",
        "Stone_Count", "Stone_Share", "Recent_Stone_Share",
    ])
    deals = ensure_columns(deals, [
        "Sold Party", "Deal_ID", "Deal_Start_Date", "Deal_End_Date",
        "Deal_Total_Stones",
    ])
    mode_summary = ensure_columns(mode_summary, ["Sold Party", "Deal_Mode", "Target_Stone_Count"])
    recipes = ensure_columns(recipes, ["Sold Party", "Deal_Mode", "Target_Quantity"])

    required_core = [
        review_book,
        output / "Selected_Deal_Stones_v6.csv",
        output / "Party_Generation_Status_v6.csv",
        output / "Recipe_Fill_Status_v6.csv",
        output / "Profile_Deal_Stone_Details_v6.csv.gz",
        output / "Combination_Profile_v6.csv.gz",
        output / "Unique_Allocation_Candidates_By_Batch_v6.zip",
    ]
    missing_core = [str(path.name) for path in required_core if not path.exists()]
    details["Missing_Core_Output_Files"] = pd.DataFrame({"Missing_File": missing_core})
    add_check(
        checks,
        "Core production output files present",
        not missing_core,
        len(missing_core),
        "Stage 3 must create the review workbook, final allocation/status exports, profile exports, and candidate archive.",
    )

    # 1. Deal IDs and bucket overlap.
    deal_id_dups = deals[deals.get("Deal_ID", pd.Series(dtype=str)).duplicated(keep=False)].copy() if not deals.empty else pd.DataFrame()
    details["Deal_ID_Duplicates"] = deal_id_dups
    add_check(checks, "Globally unique Deal IDs", deal_id_dups.empty, len(deal_id_dups), "No Deal_ID may be reused.")

    overlap_rows = []
    if not deals.empty:
        deals["Deal_Start_Date"] = pd.to_datetime(deals["Deal_Start_Date"], errors="coerce")
        deals["Deal_End_Date"] = pd.to_datetime(deals["Deal_End_Date"], errors="coerce")
        for party, g in deals.sort_values(["Sold Party", "Deal_Start_Date", "Deal_End_Date"]).groupby("Sold Party", sort=False):
            prior = None
            for _, row in g.iterrows():
                if prior is not None and pd.notna(row["Deal_Start_Date"]) and pd.notna(prior["Deal_End_Date"]) and row["Deal_Start_Date"] <= prior["Deal_End_Date"]:
                    overlap_rows.append({
                        "Sold Party": party,
                        "Previous_Deal_ID": prior.get("Deal_ID"),
                        "Previous_End": prior.get("Deal_End_Date"),
                        "Current_Deal_ID": row.get("Deal_ID"),
                        "Current_Start": row.get("Deal_Start_Date"),
                    })
                if prior is None or pd.isna(prior["Deal_End_Date"]) or row["Deal_End_Date"] > prior["Deal_End_Date"]:
                    prior = row
    overlap_df = pd.DataFrame(overlap_rows)
    details["Overlapping_Buckets"] = overlap_df
    add_check(checks, "No overlapping deal buckets", overlap_df.empty, len(overlap_df), "Deal date ranges must not overlap within a party.")

    # 2. Historical detail invariants.
    for c in ["Sold Party", "Deal_ID", "Serial_Key", "Deal_Mode", "Preferred_Combination_Key"]:
        if c in profile:
            profile[c] = profile[c].fillna("").astype(str)
    nonblank = profile[profile.get("Serial_Key", pd.Series(dtype=str)).ne("")].copy()
    same_deal_dups = nonblank[nonblank.duplicated(["Sold Party", "Deal_ID", "Serial_Key"], keep=False)].copy()
    details["Same_Deal_Serial_Duplicates"] = same_deal_dups
    add_check(checks, "Unique serial counting within each deal", same_deal_dups.empty, len(same_deal_dups), "Canonical profile detail must contain one row per Sold Party + Deal_ID + Serial No.")

    repeated_removed = int(pd.to_numeric(repeated_audit.get("Rows_Removed_From_Profile", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()) if not repeated_audit.empty else 0
    add_check(checks, "Repeated two-day serial audit retained", True, repeated_removed, "Repeated same-deal rows are preserved in the audit and excluded from economic/profile counts.", severity="WARNING")

    combo_keys = ["Sold Party", "Deal_Mode", "Preferred_Combination_Key"]
    expected = profile.groupby(combo_keys, dropna=False).size().rename("Expected_Stone_Count").reset_index()
    combo_compare = combos.merge(expected, on=combo_keys, how="outer", indicator=True)
    combo_compare["Stone_Count"] = pd.to_numeric(combo_compare.get("Stone_Count"), errors="coerce").fillna(0)
    combo_compare["Expected_Stone_Count"] = pd.to_numeric(combo_compare["Expected_Stone_Count"], errors="coerce").fillna(0)
    combo_count_diff = combo_compare[(combo_compare["_merge"] != "both") | (combo_compare["Stone_Count"] != combo_compare["Expected_Stone_Count"])].copy()
    details["Combination_Count_Diffs"] = combo_count_diff
    add_check(checks, "Per-deal serial counting across separate deals", combo_count_diff.empty, len(combo_count_diff), "Combination Stone_Count must equal per-deal serial occurrences, including valid serial reuse in separate deals.")

    # Stone-share and recent-share recomputation.
    pm = ["Sold Party", "Deal_Mode"]
    expected_total = profile.groupby(pm, dropna=False).size().rename("Expected_Total_Stones")
    expected_share = expected.merge(expected_total.reset_index(), on=pm, how="left")
    expected_share["Expected_Stone_Share"] = expected_share["Expected_Stone_Count"] / expected_share["Expected_Total_Stones"].clip(lower=1)
    share_compare = combos.merge(expected_share[combo_keys + ["Expected_Stone_Share"]], on=combo_keys, how="left")
    share_compare["Stone_Share_Diff"] = (pd.to_numeric(share_compare.get("Stone_Share"), errors="coerce") - share_compare["Expected_Stone_Share"]).abs()

    recent_n = int(float(settings.get("RECENT_DEALS_FOR_PATTERN", 5) or 5))
    dealmeta = deals[["Sold Party", "Deal_ID", "Deal_Start_Date", "Deal_Total_Stones"]].copy()
    def mode_for_stones(v: Any) -> str:
        n = int(float(v or 0))
        return "SMALL" if n <= 11 else "MEDIUM" if n <= 24 else "BULK"
    dealmeta["Deal_Mode"] = dealmeta["Deal_Total_Stones"].map(mode_for_stones)
    recent_deals = dealmeta.sort_values(["Sold Party", "Deal_Mode", "Deal_Start_Date", "Deal_ID"]).groupby(pm, group_keys=False).tail(recent_n)
    recent_keys = pd.MultiIndex.from_frame(recent_deals[["Sold Party", "Deal_Mode", "Deal_ID"]])
    profile_keys = pd.MultiIndex.from_frame(profile[["Sold Party", "Deal_Mode", "Deal_ID"]])
    recent_profile = profile[profile_keys.isin(recent_keys)].copy()
    recent_combo = recent_profile.groupby(combo_keys, dropna=False).size().rename("Expected_Recent_Count").reset_index()
    recent_total = recent_profile.groupby(pm, dropna=False).size().rename("Expected_Recent_Total").reset_index()
    recent_expected = recent_combo.merge(recent_total, on=pm, how="left")
    recent_expected["Expected_Recent_Stone_Share"] = recent_expected["Expected_Recent_Count"] / recent_expected["Expected_Recent_Total"].clip(lower=1)
    share_compare = share_compare.merge(recent_expected[combo_keys + ["Expected_Recent_Stone_Share"]], on=combo_keys, how="left")
    share_compare["Expected_Recent_Stone_Share"] = share_compare["Expected_Recent_Stone_Share"].fillna(0)
    share_compare["Recent_Stone_Share_Diff"] = (pd.to_numeric(share_compare.get("Recent_Stone_Share"), errors="coerce").fillna(0) - share_compare["Expected_Recent_Stone_Share"]).abs()
    share_diff = share_compare[(share_compare["Stone_Share_Diff"] > 1e-10) | (share_compare["Recent_Stone_Share_Diff"] > 1e-10)].copy()
    details["Combination_Share_Diffs"] = share_diff
    add_check(checks, "Historical stone-share reconciliation", share_diff.empty, len(share_diff), "Stone_Share and Recent_Stone_Share must recompute exactly from canonical per-deal detail.")

    # 3. Final allocation and pairs.
    selected["Serial_Key"] = selected["Serial_Key"].fillna("").astype(str)
    dup_alloc = selected[selected["Serial_Key"].ne("") & selected["Serial_Key"].duplicated(keep=False)].copy()
    details["Duplicate_Allocations"] = dup_alloc
    add_check(checks, "No duplicate stock allocations", dup_alloc.empty, len(dup_alloc), "A serial may appear only once in the final selected output.")

    split_pairs = pd.DataFrame()
    if "Pair_Key" in selected:
        pair_counts = selected[selected["Pair_Key"].fillna("").astype(str).ne("")].groupby("Pair_Key")["Sold Party"].nunique()
        bad_pairs = pair_counts[pair_counts.gt(1)].index
        split_pairs = selected[selected["Pair_Key"].isin(bad_pairs)].copy()
    details["Split_Pair_Keys"] = split_pairs
    add_check(checks, "No split pair keys", split_pairs.empty, split_pairs.get("Pair_Key", pd.Series(dtype=str)).nunique(), "No physical pair key may be allocated to different parties.")

    incomplete_pair_rows = []
    if "Pair_Key" in selected.columns:
        pair_selected = selected[selected["Pair_Key"].fillna("").astype(str).ne("")].copy()
        for (pair_key, party), group in pair_selected.groupby(["Pair_Key", "Sold Party"], dropna=False, sort=False):
            members_text = next((clean(x) for x in group.get("Pair_Group_Members", pd.Series(dtype=str)) if clean(x)), "")
            expected = {clean(x) for x in members_text.split("|") if clean(x)}
            if not expected:
                expected = {clean(x) for x in str(pair_key).split("_") if clean(x)}
            actual = set(group["Serial_Key"].fillna("").astype(str))
            missing = expected - actual
            if missing:
                incomplete_pair_rows.append({
                    "Pair_Key": pair_key,
                    "Sold Party": party,
                    "Expected_Serials": " | ".join(sorted(expected)),
                    "Selected_Serials": " | ".join(sorted(actual)),
                    "Missing_Serials": " | ".join(sorted(missing)),
                })
    incomplete_pairs = pd.DataFrame(incomplete_pair_rows)
    details["Incomplete_Pair_Allocations"] = incomplete_pairs
    add_check(checks, "Complete pair allocation", incomplete_pairs.empty, len(incomplete_pairs), "Every selected pair key must contain all expected pair members for the same party.")

    historical_serials = set(nonblank["Serial_Key"])
    sold_overlap = selected[selected["Serial_Key"].isin(historical_serials)].copy()
    details["Selected_Sold_History_Overlap"] = sold_overlap
    add_check(checks, "No sold-history overlap selections", sold_overlap.empty, len(sold_overlap), "Final selected serials must not overlap completed SALE history.")

    # 4. Quality domain.
    preflight = output / "Dealmaker_Data_Preflight_v6.xlsx"
    q_audit = read_sheet(preflight, "Quality_Domain_Audit") if preflight.exists() else pd.DataFrame()
    if not q_audit.empty:
        unmapped_mask = q_audit.astype(str).apply(lambda col: col.str.upper().eq("UNMAPPED")).any(axis=1)
        unmapped = q_audit[unmapped_mask].copy()
    else:
        unmapped = pd.DataFrame()
    details["Unmapped_Quality_Values"] = unmapped
    add_check(checks, "Zero unmapped quality values", unmapped.empty, len(unmapped), "Any nonblank UNMAPPED quality value blocks production.")

    # 5. Recipe fill final reconciliation.
    required_line_cols = {
        "Requested_Recipe_Quantity", "Candidate_Availability", "Pre_Allocation_Filled_Quantity",
        "Final_Globally_Allocated_Quantity", "Missing_Quantity", "Failure_Reason",
        "Recipe_Fill_Reconciliation_Status",
    }
    missing_cols = sorted(required_line_cols - set(line.columns))
    final_counts = selected.groupby(["Sold Party", "Deal_Mode", "Recipe_Line_ID"], dropna=False).size().rename("Expected_Final").reset_index()
    line_compare = line.merge(final_counts, on=["Sold Party", "Deal_Mode", "Recipe_Line_ID"], how="left")
    line_compare["Expected_Final"] = line_compare["Expected_Final"].fillna(0).astype(int)
    line_compare["Final_Globally_Allocated_Quantity"] = pd.to_numeric(line_compare.get("Final_Globally_Allocated_Quantity"), errors="coerce").fillna(-1).astype(int)
    line_bad = line_compare[
        (line_compare["Final_Globally_Allocated_Quantity"] != line_compare["Expected_Final"])
        | line_compare.get("Recipe_Fill_Reconciliation_Status", pd.Series("MISSING", index=line_compare.index)).ne("RECONCILED")
    ].copy()
    details["Recipe_Fill_Diffs"] = line_bad
    recipe_ok = not missing_cols and line_bad.empty and int(line_compare["Final_Globally_Allocated_Quantity"].sum()) == len(selected)
    add_check(checks, "Recipe-fill reconciliation", recipe_ok, len(line_bad), f"Required columns missing: {missing_cols or 'none'}; final line quantities must sum to selected stones.")

    # 6. Current action windows.
    action_bad = []
    for _, row in selected.iterrows():
        action = clean(row.get("Primary_Action_Type")).upper()
        if not action or action in {"NAN", "NONE"}:
            continue
        age = pd.to_numeric(pd.Series([row.get("Action_Age_Days")]), errors="coerce").iloc[0]
        max_age = pd.to_numeric(pd.Series([settings.get(f"{action}_MAX_AGE_DAYS")]), errors="coerce").iloc[0]
        if pd.notna(age) and pd.notna(max_age) and float(age) > float(max_age):
            action_bad.append({"Sold Party": row.get("Sold Party"), "Serial_Key": row.get("Serial_Key"), "Action": action, "Action_Age_Days": age, "Maximum_Days": max_age})
    action_bad_df = pd.DataFrame(action_bad)
    details["Action_Window_Failures"] = action_bad_df
    add_check(checks, "Action-window compliance", action_bad_df.empty, len(action_bad_df), "Action-linked selected stones must be within the configured recency window.")

    # 7. Live-stock release blocking.
    status_present = False
    if not data_quality.empty and {"Check", "Value"}.issubset(data_quality.columns):
        row = data_quality[data_quality["Check"].astype(str).eq("Explicit availability/status column populated")]
        if not row.empty:
            status_present = clean(row.iloc[-1]["Value"]).upper() == "YES"
    snapshot_live = yes(settings.get("STOCK_SNAPSHOT_CONFIRMED_LIVE", "NO"))
    generated = status[status.get("Generation_Status", pd.Series(dtype=str)).astype(str).eq("GENERATED")].copy()
    if not status_present and not snapshot_live:
        live_bad = generated[~generated.get("Release_Status", pd.Series(index=generated.index, dtype=str)).astype(str).str.contains("LIVE STOCK CONFIRMATION REQUIRED", na=False)].copy()
    else:
        live_bad = pd.DataFrame()
    details["Live_Stock_Blocking_Failures"] = live_bad
    add_check(checks, "Live-stock release blocking", live_bad.empty, len(live_bad), "Without a populated status field or explicit live snapshot confirmation, every generated basket must remain blocked.")

    # 8. Salesperson and combined workbooks.
    selected_keys = selected[["Sold Party", "Serial_Key"]].fillna("").astype(str).drop_duplicates()
    salesperson_rows = []
    salesperson_link_failures = []
    salesperson_link_metrics = []
    formula_errors = 0
    sp_zip = output / "Generated_Deal_Baskets_By_Salesperson_UNIQUE_v6.zip"
    if sp_zip.exists():
        with zipfile.ZipFile(sp_zip) as zf:
            for name in [n for n in zf.namelist() if n.lower().endswith(".xlsx")]:
                data = zf.read(name)
                part = serial_rows_from_workbook_bytes(data)
                if not part.empty:
                    part["Workbook"] = name
                    salesperson_rows.append(part)
                link_bad, link_metrics = native_link_metrics_from_workbook_bytes(data, name)
                salesperson_link_metrics.append(link_metrics)
                if not link_bad.empty:
                    salesperson_link_failures.append(link_bad)
                formula_errors += check_formula_errors_workbook(data)
    elif not selected.empty:
        salesperson_link_failures.append(pd.DataFrame([{
            "Workbook": sp_zip.name, "Sheet": "", "Cell": "", "Header": "ALL",
            "Value": "", "Target_Length": 0, "Reason": "SALESPERSON_ZIP_MISSING",
        }]))
    salesperson = pd.concat(salesperson_rows, ignore_index=True) if salesperson_rows else pd.DataFrame(columns=["Sold Party", "Serial_Key"])
    sp_keys = salesperson[["Sold Party", "Serial_Key"]].fillna("").astype(str).drop_duplicates()
    sp_diff = selected_keys.merge(sp_keys, on=["Sold Party", "Serial_Key"], how="outer", indicator=True)
    sp_diff = sp_diff[sp_diff["_merge"].ne("both")].copy()
    salesperson_link_bad = pd.concat(salesperson_link_failures, ignore_index=True) if salesperson_link_failures else pd.DataFrame()
    salesperson_link_metrics_df = pd.DataFrame(salesperson_link_metrics)
    expected_links = expected_native_link_count(selected)
    actual_links = int(pd.to_numeric(salesperson_link_metrics_df.get("Native_HTTP_Links", pd.Series(dtype=int)), errors="coerce").fillna(0).sum())
    formula_links = int(pd.to_numeric(salesperson_link_metrics_df.get("Formula_Hyperlinks", pd.Series(dtype=int)), errors="coerce").fillna(0).sum())
    xml_formula_links = int(pd.to_numeric(salesperson_link_metrics_df.get("Worksheet_XML_Hyperlink_Formulas", pd.Series(dtype=int)), errors="coerce").fillna(0).sum())
    if actual_links != expected_links or formula_links or xml_formula_links:
        count_failure = pd.DataFrame([{
            "Workbook": "ALL SALESPERSON WORKBOOKS",
            "Sheet": "",
            "Cell": "",
            "Header": "ALL",
            "Value": f"native={actual_links}; expected={expected_links}; formulas={formula_links}; xml_formulas={xml_formula_links}",
            "Target_Length": 0,
            "Reason": "NATIVE_LINK_COUNT_OR_FORMULA_GATE_FAILED",
        }])
        salesperson_link_bad = pd.concat([salesperson_link_bad, count_failure], ignore_index=True)
    details["Salesperson_Output_Diffs"] = sp_diff
    details["Salesperson_Link_Failures"] = salesperson_link_bad
    details["Salesperson_Link_Metrics"] = salesperson_link_metrics_df
    salesperson_ok = sp_diff.empty and salesperson_link_bad.empty
    add_check(
        checks,
        "Salesperson-output reconciliation",
        salesperson_ok,
        len(sp_diff) + len(salesperson_link_bad),
        "Salesperson workbooks must contain exactly the final selected party/serial rows and use valid native text hyperlinks for URL-backed stock fields.",
    )

    combined_path = output / "Generated_Deal_Baskets_UNIQUE_v6.xlsx"
    if combined_path.exists():
        combined_bytes = combined_path.read_bytes()
        combined = serial_rows_from_workbook_bytes(combined_bytes)
        combined_keys = combined[["Sold Party", "Serial_Key"]].fillna("").astype(str).drop_duplicates() if not combined.empty else pd.DataFrame(columns=["Sold Party", "Serial_Key"])
        combined_diff = selected_keys.merge(combined_keys, on=["Sold Party", "Serial_Key"], how="outer", indicator=True)
        combined_diff = combined_diff[combined_diff["_merge"].ne("both")].copy()
        combined_link_bad, combined_link_metrics = native_link_metrics_from_workbook_bytes(combined_bytes, combined_path.name)
    else:
        combined_diff = selected_keys.copy() if not selected.empty else pd.DataFrame(columns=["Sold Party", "Serial_Key"])
        combined_link_bad = pd.DataFrame()
        combined_link_metrics = {
            "Workbook": combined_path.name, "Native_HTTP_Links": 0,
            "Formula_Hyperlinks": 0, "Visible_Raw_URLs": 0,
            "Invalid_Native_Targets": 0, "Worksheet_XML_Hyperlink_Formulas": 0,
        }
        if not selected.empty:
            combined_link_bad = pd.DataFrame([{
                "Workbook": combined_path.name, "Sheet": "", "Cell": "", "Header": "ALL",
                "Value": "", "Target_Length": 0, "Reason": "COMBINED_WORKBOOK_MISSING",
            }])
    combined_actual_links = int(combined_link_metrics["Native_HTTP_Links"])
    combined_formula_links = int(combined_link_metrics["Formula_Hyperlinks"])
    combined_xml_formula_links = int(combined_link_metrics["Worksheet_XML_Hyperlink_Formulas"])
    if combined_actual_links != expected_links or combined_formula_links or combined_xml_formula_links:
        count_failure = pd.DataFrame([{
            "Workbook": combined_path.name,
            "Sheet": "",
            "Cell": "",
            "Header": "ALL",
            "Value": f"native={combined_actual_links}; expected={expected_links}; formulas={combined_formula_links}; xml_formulas={combined_xml_formula_links}",
            "Target_Length": 0,
            "Reason": "NATIVE_LINK_COUNT_OR_FORMULA_GATE_FAILED",
        }])
        combined_link_bad = pd.concat([combined_link_bad, count_failure], ignore_index=True)
    details["Combined_Output_Diffs"] = combined_diff
    details["Combined_Link_Failures"] = combined_link_bad
    details["Combined_Link_Metrics"] = pd.DataFrame([combined_link_metrics])
    combined_ok = combined_diff.empty and combined_link_bad.empty
    add_check(
        checks,
        "Combined-basket reconciliation",
        combined_ok,
        len(combined_diff) + len(combined_link_bad),
        "Combined workbook must contain exactly the final selected party/serial rows and valid native text hyperlinks.",
    )

    if review_book.exists():
        formula_errors += check_formula_errors_workbook(review_book)
    if combined_path.exists():
        formula_errors += check_formula_errors_workbook(combined_path)
    add_check(checks, "No Excel formula errors", formula_errors == 0, formula_errors, "Review, combined, and salesperson workbooks must contain no explicit Excel error values.")

    # 9. Candidate archive semantics and final flags.
    candidate_bad = []
    candidate_final_keys: set[tuple[str, str, str, str]] = set()
    required_candidate_cols = {
        "Pre_Allocation_Candidate_Status", "Candidate_Archive_Stage",
        "Candidate_Archive_Is_Final_Output", "Final_Globally_Allocated_Flag",
        "Final_Allocation_Status",
    }
    candidate_zip = output / "Unique_Allocation_Candidates_By_Batch_v6.zip"
    semantics_present = False
    if candidate_zip.exists():
        with zipfile.ZipFile(candidate_zip) as zf:
            semantics_present = "CANDIDATE_ARCHIVE_SEMANTICS_v6.txt" in zf.namelist()
            for name in [n for n in zf.namelist() if n.endswith(".csv.gz")]:
                raw = zf.read(name)
                with gzip.GzipFile(fileobj=io.BytesIO(raw)) as fh:
                    batch = pd.read_csv(fh, low_memory=False)
                missing = required_candidate_cols - set(batch.columns)
                if missing:
                    candidate_bad.append({"File": name, "Issue": f"Missing columns: {sorted(missing)}"})
                    continue
                if not batch["Candidate_Archive_Is_Final_Output"].astype(str).eq("NO").all():
                    candidate_bad.append({"File": name, "Issue": "Archive incorrectly marked as final output"})
                final_rows = batch[batch["Final_Globally_Allocated_Flag"].astype(str).eq("YES")]
                for values in final_rows[["Sold Party", "Deal_Mode", "Recipe_Line_ID", "Serial_Key"]].fillna("").astype(str).itertuples(index=False, name=None):
                    candidate_final_keys.add(tuple(values))
    else:
        candidate_bad.append({"File": candidate_zip.name, "Issue": "Candidate archive missing"})
    selected_candidate_keys = set(map(tuple, selected[["Sold Party", "Deal_Mode", "Recipe_Line_ID", "Serial_Key"]].fillna("").astype(str).itertuples(index=False, name=None)))
    if candidate_final_keys != selected_candidate_keys:
        candidate_bad.append({"File": "ALL", "Issue": f"Final-flag key difference: archive={len(candidate_final_keys)}, selected={len(selected_candidate_keys)}"})
    if not semantics_present:
        candidate_bad.append({"File": "ALL", "Issue": "Semantics text missing"})
    candidate_bad_df = pd.DataFrame(candidate_bad)
    details["Candidate_Archive_Failures"] = candidate_bad_df
    add_check(checks, "Candidate-archive semantics", candidate_bad_df.empty, len(candidate_bad_df), "Archive must be explicitly non-final and final-allocation flags must reconcile to Selected_Deal_Stones.")

    # 10. Recipe target must equal the historical mode target before allocation.
    recipe_target_bad = pd.DataFrame()
    if not recipes.empty and not mode_summary.empty:
        recipe_totals = recipes.groupby(["Sold Party", "Deal_Mode"], dropna=False)["Target_Quantity"].sum().rename("Recipe_Target_Total").reset_index()
        expected = mode_summary[["Sold Party", "Deal_Mode", "Target_Stone_Count"]].copy()
        recipe_target_bad = expected.merge(recipe_totals, on=["Sold Party", "Deal_Mode"], how="left")
        recipe_target_bad["Recipe_Target_Total"] = pd.to_numeric(recipe_target_bad["Recipe_Target_Total"], errors="coerce").fillna(0).astype(int)
        recipe_target_bad["Target_Stone_Count"] = pd.to_numeric(recipe_target_bad["Target_Stone_Count"], errors="coerce").fillna(0).astype(int)
        recipe_target_bad = recipe_target_bad[recipe_target_bad["Recipe_Target_Total"].ne(recipe_target_bad["Target_Stone_Count"])].copy()
    details["Recipe_Target_Mismatches"] = recipe_target_bad
    add_check(checks, "Recipe target equals historical mode target", recipe_target_bad.empty, len(recipe_target_bad), "Every party/mode recipe must sum exactly to Party_Deal_Mode_Summary.Target_Stone_Count.")

    # 11. Accepted basket sizes must comply with their mode labels.
    generated_status = status[status.get("Generation_Status", pd.Series(dtype=str)).astype(str).eq("GENERATED")].copy()
    size_bad_rows = []
    min_size = {"SMALL": 6, "MEDIUM": 12, "BULK": 25}
    max_size = {"SMALL": 11, "MEDIUM": 24, "BULK": None}
    for _, row in generated_status.iterrows():
        mode = clean(row.get("Chosen_Deal_Mode", "")).upper()
        count = int(pd.to_numeric(pd.Series([row.get("Normal_Selected_Stone_Count", row.get("Selected_Stone_Count", 0))]), errors="coerce").fillna(0).iloc[0])
        final_count = int(pd.to_numeric(pd.Series([row.get("Selected_Stone_Count", count)]), errors="coerce").fillna(count).iloc[0])
        pair_added = int(pd.to_numeric(pd.Series([row.get("Pair_Added_Stone_Count", max(final_count - count, 0))]), errors="coerce").fillna(0).iloc[0])
        low = min_size.get(mode, 0); high = max_size.get(mode)
        if count < low or (high is not None and count > high):
            size_bad_rows.append({"Sold Party": row.get("Sold Party", ""), "Deal_Mode": mode, "Normal_Selected_Stone_Count": count, "Pair_Added_Stone_Count": pair_added, "Final_Selected_Stone_Count": final_count, "Minimum": low, "Maximum": high})
    size_bad = pd.DataFrame(size_bad_rows)
    details["Generated_Basket_Size_Failures"] = size_bad
    add_check(checks, "Accepted basket mode-size compliance", size_bad.empty, len(size_bad), "Normal ranked selection must satisfy SMALL 6-11, MEDIUM 12-24, and BULK 25+; mandatory pair additions may increase the final basket size.")

    # 12. The chosen mode must be due under its own clock unless explicitly overridden.
    allowed_status = {"DUE NOW", "OVERDUE"} | ({"DUE SOON"} if yes(settings.get("INCLUDE_DUE_SOON", "NO")) else set())
    if not generated_status.empty:
        override = generated_status.get("Mode_Selection_Override", pd.Series("NO", index=generated_status.index)).astype(str).str.upper()
        chosen_status = generated_status.get("Chosen_Mode_Interval_Status", pd.Series("", index=generated_status.index)).astype(str).str.upper()
        mode_due_bad = generated_status[~chosen_status.isin(allowed_status) & ~override.isin({"SMALL", "MEDIUM", "BULK"})].copy()
    else:
        mode_due_bad = pd.DataFrame()
    details["Chosen_Mode_Cadence_Failures"] = mode_due_bad
    add_check(checks, "Chosen mode has due independent clock", mode_due_bad.empty, len(mode_due_bad), "Automatic baskets may use only modes that are due under their own cadence; manual forced-mode overrides are disclosed.")

    # 13. Actions are boost-only by default: they can label/rank, but cannot reject a due valid basket.
    policy = clean(settings.get("ACTION_POLICY", "BOOST_ONLY")).upper() or "BOOST_ONLY"
    if policy == "BOOST_ONLY":
        invalid_labels = generated_status[
            ~generated_status.get("Action_Support_Status", pd.Series("", index=generated_status.index)).astype(str).isin({"ACTION SUPPORTED", "SALES HISTORY ONLY"})
            | ~generated_status.get("Action_Data_Status", pd.Series("", index=generated_status.index)).astype(str).isin({"CURRENT ACTION DATA", "NO RECENT ACTION DATA", "STALE ACTION DATA"})
        ].copy() if not generated_status.empty else pd.DataFrame()
        action_rejection = status[
            status.get("Generation_Reason", pd.Series("", index=status.index)).astype(str).str.contains("action gate not met", case=False, na=False)
        ].copy() if not status.empty else pd.DataFrame()
        action_policy_bad = pd.concat([invalid_labels, action_rejection], ignore_index=True) if not invalid_labels.empty or not action_rejection.empty else pd.DataFrame()
        detail = "Under BOOST_ONLY, Actions may rank candidates and create review labels, but may not reject an otherwise valid due basket."
    else:
        action_policy_bad = generated_status[generated_status.get("Mode_Action_Gate_Passed", pd.Series("YES", index=generated_status.index)).astype(str).str.upper().ne("YES")].copy() if not generated_status.empty else pd.DataFrame()
        detail = "Under HARD_GATE, generated MEDIUM/BULK baskets must meet configured action thresholds."
    details["Action_Policy_Failures"] = action_policy_bad
    add_check(checks, "Action-policy compliance", action_policy_bad.empty, len(action_policy_bad), detail)

    # 14. Chosen recipe lines must have resolved quality limits and disclose fallback source.
    quality_bad = pd.DataFrame()
    if not recipes.empty and not generated_status.empty:
        chosen_keys = generated_status[["Sold Party", "Chosen_Deal_Mode"]].rename(columns={"Chosen_Deal_Mode": "Deal_Mode"}).drop_duplicates()
        chosen_recipes = recipes.merge(chosen_keys, on=["Sold Party", "Deal_Mode"], how="inner")
        limit_cols = [c for c in ["Normal_Max_Shade", "Normal_Max_Milky", "Normal_Max_Open", "Normal_Max_Blk"] if c in chosen_recipes.columns]
        if limit_cols:
            bad_mask = chosen_recipes[limit_cols].fillna("").astype(str).apply(lambda c: c.str.strip().str.upper().isin({"", "UNKNOWN", "NAN"})).any(axis=1)
            quality_bad = chosen_recipes[bad_mask].copy()
    details["Chosen_Recipe_Quality_Limit_Failures"] = quality_bad
    add_check(checks, "Chosen recipe quality-limit completeness", quality_bad.empty, len(quality_bad), "Every chosen recipe line must have resolved Shade, Milky, Open, and black-inclusion limits with an auditable source.")

    return pd.DataFrame(checks), details


def main() -> None:
    ap = argparse.ArgumentParser(description="Independent Dealmaker v8.2.2 release-gate audit.")
    ap.add_argument("--output-dir", required=True)
    ap.add_argument("--control", required=True)
    ap.add_argument("--audit-workbook", default="")
    ap.add_argument("--summary", default="")
    args = ap.parse_args()
    output = Path(args.output_dir).resolve()
    control = Path(args.control).resolve()
    checks, details = run(output, control)
    audit_path = Path(args.audit_workbook).resolve() if args.audit_workbook else output / "Independent_Release_Gate_Audit_v8.xlsx"
    summary_path = Path(args.summary).resolve() if args.summary else output / "RELEASE_GATE_SUMMARY_v8.txt"
    with pd.ExcelWriter(audit_path, engine="openpyxl") as writer:
        checks.to_excel(writer, sheet_name="Release_Gates", index=False)
        for name, df in details.items():
            (df if not df.empty else pd.DataFrame({"Result": ["No exceptions"]})).to_excel(writer, sheet_name=name[:31], index=False)
    style_workbook(audit_path)
    failures = checks[checks["Status"].eq("FAIL")]
    warnings = checks[checks["Status"].eq("WARNING")]
    with summary_path.open("w", encoding="utf-8") as fh:
        fh.write(f"DEALMAKER INDEPENDENT RELEASE-GATE SUMMARY v{VERSION}\n")
        fh.write("=" * 54 + "\n")
        fh.write(f"Overall result: {'PASS' if failures.empty else 'FAIL'}\n")
        fh.write(f"Passed gates: {int(checks['Status'].eq('PASS').sum())}\n")
        fh.write(f"Warnings: {len(warnings)}\n")
        fh.write(f"Failed gates: {len(failures)}\n\n")
        for row in checks.itertuples(index=False):
            fh.write(f"[{row.Status}] {row.Gate}: {row.Value} - {row.Detail}\n")
        if failures.empty:
            fh.write("\nMechanical release gates passed. Commercial release remains governed by each party's Release_Status, including live-stock and freshness blocks.\n")
    print(checks.to_string(index=False))
    print(f"\nAudit workbook: {audit_path}\nSummary: {summary_path}")
    raise SystemExit(0 if failures.empty else 1)


if __name__ == "__main__":
    main()
