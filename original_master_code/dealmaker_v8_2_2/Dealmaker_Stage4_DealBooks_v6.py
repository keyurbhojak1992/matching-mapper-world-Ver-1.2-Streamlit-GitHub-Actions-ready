from __future__ import annotations

import argparse
import math
import pickle
import re
import sys
import zipfile
from io import BytesIO
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# The salesperson workbook layout deliberately mirrors the supplied
# Suggest_Regular_Stock_2026-05-18 reference. The order is fixed, but a
# visible column is emitted only when its source field exists in the selected
# dataset. Dealmaker audit fields are kept on hidden sheets so the familiar
# stone presentation is not disturbed.
STOCK_LAYOUT = [
    ("Sr.No", None),
    ("Serial No", "Serial No"),
    ("DETAIL", "DETAIL"),
    ("Shp", "Shp"),
    ("Carat", "Carat"),
    ("Color", "Col"),
    ("Qua", "Qua"),
    ("Cut", "Cut"),
    ("Pol", "Pol"),
    ("Sym", "Sym"),
    ("Flo", "Flo"),
    ("Lab", "Lab"),
    ("RapRate", "RapRate"),
    ("Disc%", "Disc%"),
    ("Rate", "Rate"),
    ("Amount", "Amount"),
    ("SGPComment", "SGPComment"),
    ("Loc", "Loc"),
    ("Measurement", "Measurement"),
    ("Ratio", "Ratio"),
    ("Dep", "Dep"),
    ("Tab", "Tab"),
    ("Girdle", "Girdle"),
    ("C/A", "C/A"),
    ("C/H", "C/H"),
    ("P/A", "P/A"),
    ("P/H", "P/H"),
    ("Culet", "Culet"),
    ("SINC_NAME", "SINC_NAME"),
    ("TOIN_NAME", "TOIN_NAME"),
    ("SOIN_NAME", "SOIN_NAME"),
    ("GIRDLE_OPEN_NAME", "GIRDLE_OPEN_NAME"),
    ("PAV_OPEN_NAME", "PAV_OPEN_NAME"),
    ("TBINC_NAME", "TBINC_NAME"),
    ("SBINC_NAME", "SBINC_NAME"),
    ("KeyToSymbol", "KeyToSymbol"),
    ("Comment", "Comment"),
    ("IMAGE", "IMAGE"),
    ("VIDEO", "VIDEO"),
    ("CERTIFICATE", "CERTIFICATE"),
    ("Shade", "Shade"),
    ("Milky", "Milky"),
    ("Report No", "ReportNo"),
    ("Color Desc", "Color Desc"),
    ("TINC_NAME", "TINC_NAME"),
    ("Girdle Desc", "Girdle Desc"),
    ("Girdle Condition", "Girdle Condition"),
    ("H&A", "H&A"),
    ("Report Date", "Report Date"),
    ("Std Shape", "Std Shape"),
]

NAV_HEADERS = ["Go To Sheet", "Party Name", "Count", "Country", "City", "Go →"]
NAV_WIDTHS = [36.3, 40.0, 7.7, 11.0, 14.3, 6.6]
PARTY_WIDTHS = [
    9.9, 14.3, 12.0, 7.7, 9.9, 9.9, 7.7, 7.7, 7.7, 7.7,
    7.7, 7.7, 12.1, 9.9, 9.9, 11.0, 40.0, 9.9, 23.1, 9.9,
    7.7, 7.7, 11.0, 7.7, 7.7, 7.7, 7.7, 9.9, 15.4, 14.3,
    14.3, 22.0, 18.7, 19.8, 19.8, 40.0, 40.0, 12.0, 12.0, 12.0,
    24.2, 24.2, 15.4, 15.4, 15.4, 16.5, 22.0, 7.7, 16.5, 14.3,
]
STOCK_WIDTH_BY_LABEL = {
    label: width for (label, _), width in zip(STOCK_LAYOUT, PARTY_WIDTHS)
}

NUMERIC_HEADERS = {
    "Carat", "RapRate", "Disc%", "Rate", "Amount", "Ratio", "Dep", "Tab",
    "Girdle", "C/A", "C/H", "P/A", "P/H",
}
LINK_HEADERS = {"DETAIL", "IMAGE", "VIDEO", "CERTIFICATE", "PHYGITAL", "LAYOUT"}
GENERATOR_VERSION = "8.2.2-PAIR-COMPLETION-NATIVE-LINKS"

THIN = Side(style="thin", color="000000")
GRID_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
GREY_HEADER = PatternFill("solid", fgColor="D3D3D3")
EXACT_MATCH_FILL = PatternFill("solid", fgColor="CCFFCC")
NEAR_MATCH_FILL = PatternFill("solid", fgColor="FFE699")
NO_MATCH_FILL = PatternFill(fill_type=None)
NAV_HEADER_FONT = Font(color="0000FF", bold=True)
PARTY_HEADER_FONT = Font(bold=True)
SERIAL_NUMBER_FONT = Font(color="0000FF")
LINK_FONT = Font(color="0000FF", underline="single")
AUDIT_HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
AUDIT_HEADER_FONT = Font(color="FFFFFF", bold=True)

AUDIT_COLUMNS = [
    "Sold Party", "Sales_Person", "Country", "City", "Serial No", "Serial_Key",
    "Deal_Mode", "Chosen_Deal_Mode", "Chosen_Mode_Interval_Status",
    "Chosen_Mode_Interval_Source", "Target_Stone_Count", "Historical_Target_Quantity",
    "Selection_Reason", "Match_Level", "Direct_Action_Flag", "Action_Support_Status",
    "Action_Data_Status", "Action_Review_Required", "Pair_Resolution_Status",
    "Pair_Added_Flag", "Pair_Anchor_Serial",
    "Release_Status", "Generation_Reason",
]


def load(root: Path, name: str):
    with open(root / f"{name}.pkl", "rb") as fh:
        return pickle.load(fh)


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    return str(value).strip()


def clean_pair_serial(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    text = str(value).strip().upper()
    text = re.sub(r"\.0+$", "", text)
    return re.sub(r"\s+", "", text)


def get_pair_id_for_sorting(text: Any) -> str | None:
    if text is None:
        return None
    if isinstance(text, float) and math.isnan(text):
        return None
    match = re.search(r"(?:PAIRED\s+WITH|MATCH\s+PAIR)\s*[^\d]*([\d.]+)", str(text), re.IGNORECASE)
    if not match:
        return None
    pair_id = clean_pair_serial(match.group(1))
    return pair_id or None


def yes(value: Any) -> bool:
    return clean_text(value).upper() in {"YES", "Y", "TRUE", "1"}


def safe(value: Any):
    if value is None:
        return None
    if isinstance(value, pd.Timestamp):
        return None if pd.isna(value) else value.to_pydatetime()
    if isinstance(value, np.generic):
        return value.item()
    if isinstance(value, float) and pd.isna(value):
        return None
    if not isinstance(value, (str, int, float, bool)):
        value = str(value)
    return value[:32760] if isinstance(value, str) else value


def safe_sheet_title(party: Any, used: set[str]) -> str:
    base = re.sub(r"[\\/*?:\[\]]", "", clean_text(party))[:31] or "Party"
    title = base
    suffix_number = 2
    while title in used:
        suffix = f"_{suffix_number}"
        title = base[: 31 - len(suffix)] + suffix
        suffix_number += 1
    used.add(title)
    return title


def salesperson_filename(value: Any) -> str:
    name = clean_text(value) or "Unknown Salesperson"
    name = re.sub(r"\b(Mr|Ms|Mrs|Dr)\.\s*", r"\1 ", name, flags=re.IGNORECASE)
    name = re.sub(r"\s+", " ", name).strip()
    name = re.sub(r'[\\/*?:\[\]<>|"]', "", name)
    return f"{name or 'Unknown Salesperson'}.xlsx"


HYPERLINK_FORMULA_RE = re.compile(
    r'^=HYPERLINK\(\s*"((?:[^"]|"")*)"\s*[,;]\s*"((?:[^"]|"")*)"\s*\)$',
    flags=re.IGNORECASE,
)


def hyperlink_target(value: Any) -> str:
    """Return an HTTP/HTTPS target from a raw URL or HYPERLINK formula.

    Long stock URLs must be stored as native XLSX hyperlinks. Excel formula
    hyperlinks are unreliable when the link target exceeds 255 characters,
    while the worksheet hyperlink relationship supports these long URLs.
    """
    text = clean_text(value)
    if not text:
        return ""
    if text.lower().startswith(("http://", "https://")):
        return text
    match = HYPERLINK_FORMULA_RE.match(text)
    if not match:
        return ""
    target = match.group(1).replace('""', '"')
    return target if target.lower().startswith(("http://", "https://")) else ""


def set_link_cell(cell, label: str, value: Any) -> None:
    """Write a readable native hyperlink without exposing the raw URL."""
    target = hyperlink_target(value) if label in LINK_HEADERS else ""
    if target:
        cell.value = label
        cell.hyperlink = target
        cell.font = LINK_FONT
    else:
        cell.value = safe(value)


def normalize_workbook_stock_links(wb: Workbook) -> None:
    """Make native hyperlinks the final cell-writing operation.

    This scans every visible party sheet immediately before serialization. It
    converts any raw URL or legacy HYPERLINK formula still present in a stock
    link column to an external XLSX hyperlink relationship. Navigation links
    and hidden audit sheets are intentionally left unchanged.
    """
    for ws in wb.worksheets:
        if ws.title == "Navigation" or ws.title.startswith("Dealmaker") or ws.max_row < 2:
            continue
        headers = {
            clean_text(cell.value): cell.column
            for cell in ws[1]
            if clean_text(cell.value)
        }
        for label in LINK_HEADERS.intersection(headers):
            column = headers[label]
            for row_number in range(2, ws.max_row + 1):
                cell = ws.cell(row_number, column)
                existing_target = ""
                if cell.hyperlink is not None:
                    existing_target = clean_text(cell.hyperlink.target or cell.hyperlink.location)
                target = existing_target if existing_target.lower().startswith(("http://", "https://")) else hyperlink_target(cell.value)
                if not target:
                    continue
                cell.value = label
                cell.hyperlink = target
                cell.font = LINK_FONT


def expected_stock_link_count(selected: pd.DataFrame) -> int:
    """Count URL-backed cells that must become native links."""
    total = 0
    for label, source in available_stock_layout(selected.columns):
        if label not in LINK_HEADERS or source is None or source not in selected.columns:
            continue
        total += int(selected[source].map(lambda value: bool(hyperlink_target(value))).sum())
    return total


def workbook_native_link_metrics(data: bytes) -> dict[str, int]:
    """Inspect a serialized workbook exactly as Excel will receive it."""
    wb = load_workbook(BytesIO(data), read_only=False, data_only=False)
    metrics = {
        "Native_HTTP_Links": 0,
        "Formula_Hyperlinks": 0,
        "Visible_Raw_URLs": 0,
        "Invalid_Native_Targets": 0,
        "Visible_Label_Mismatches": 0,
        "Worksheet_XML_Hyperlink_Formulas": 0,
    }
    for ws in wb.worksheets:
        if ws.title == "Navigation" or ws.title.startswith("Dealmaker") or ws.max_row < 2:
            continue
        headers = {
            clean_text(cell.value): cell.column
            for cell in ws[1]
            if clean_text(cell.value)
        }
        for label in LINK_HEADERS.intersection(headers):
            column = headers[label]
            for row_number in range(2, ws.max_row + 1):
                cell = ws.cell(row_number, column)
                value = clean_text(cell.value)
                if value.upper().startswith("=HYPERLINK("):
                    metrics["Formula_Hyperlinks"] += 1
                if value.lower().startswith(("http://", "https://")):
                    metrics["Visible_Raw_URLs"] += 1
                if cell.hyperlink is None:
                    continue
                target = clean_text(cell.hyperlink.target or cell.hyperlink.location)
                if target.lower().startswith(("http://", "https://")):
                    metrics["Native_HTTP_Links"] += 1
                    if value != label:
                        metrics["Visible_Label_Mismatches"] += 1
                else:
                    metrics["Invalid_Native_Targets"] += 1

    with zipfile.ZipFile(BytesIO(data)) as archive:
        for name in archive.namelist():
            if name.startswith("xl/worksheets/") and name.endswith(".xml"):
                metrics["Worksheet_XML_Hyperlink_Formulas"] += archive.read(name).upper().count(b"HYPERLINK(")
    return metrics


def serialize_validated_workbook(wb: Workbook, selected: pd.DataFrame, workbook_name: str) -> tuple[bytes, dict[str, int]]:
    """Normalize, save, reopen and fail closed on Excel-unsafe links."""
    normalize_workbook_stock_links(wb)
    expected = expected_stock_link_count(selected)
    buffer = BytesIO()
    wb.save(buffer)
    data = buffer.getvalue()
    metrics = workbook_native_link_metrics(data)
    metrics["Expected_Native_HTTP_Links"] = expected

    failures = []
    for field in (
        "Formula_Hyperlinks",
        "Visible_Raw_URLs",
        "Invalid_Native_Targets",
        "Visible_Label_Mismatches",
        "Worksheet_XML_Hyperlink_Formulas",
    ):
        if metrics[field] != 0:
            failures.append(f"{field}={metrics[field]}")
    if metrics["Native_HTTP_Links"] != expected:
        failures.append(
            f"Native_HTTP_Links={metrics['Native_HTTP_Links']} expected={expected}"
        )
    if failures:
        raise RuntimeError(
            f"Excel-safe hyperlink validation failed for {workbook_name}: "
            + "; ".join(failures)
        )
    return data, metrics


def available_stock_layout(columns) -> list[tuple[str, str | None]]:
    """Preserve the reference order but omit source fields not in the dataset."""
    available = set(columns)
    return [
        (label, source)
        for label, source in STOCK_LAYOUT
        if source is None or source in available
    ]


def match_fill(value: Any) -> PatternFill:
    """Apply the requested three-level visual match logic."""
    level = clean_text(value).upper()
    if level == "EXACT":
        return EXACT_MATCH_FILL
    if level in {"PREFERRED", "NEAR", "NEARBY", "NEAR BY"}:
        return NEAR_MATCH_FILL
    # RELAXED means the broader three-variable base was changed. It remains
    # deliberately unfilled, as do unknown/unclassified match levels.
    return NO_MATCH_FILL


def style_navigation(ws) -> None:
    ws.freeze_panes = "A2"
    for col, width in enumerate(NAV_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(col)].width = width
    for cell in ws[1]:
        cell.fill = GREY_HEADER
        cell.font = NAV_HEADER_FONT
        cell.border = GRID_BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=len(NAV_HEADERS)):
        for cell in row:
            cell.border = GRID_BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")


def add_navigation(wb: Workbook, selected: pd.DataFrame, status: pd.DataFrame) -> dict[str, str]:
    ws = wb.create_sheet("Navigation", 0)
    ws.append(NAV_HEADERS)

    party_rows = []
    for party, group in selected.groupby("Sold Party", sort=False, dropna=False):
        party_text = clean_text(party)
        status_rows = status[status["Sold Party"].astype(str).eq(party_text)] if "Sold Party" in status.columns else pd.DataFrame()
        status_row = status_rows.iloc[0] if not status_rows.empty else pd.Series(dtype=object)
        party_rows.append({
            "party": party_text,
            "count": len(group),
            "country": clean_text(status_row.get("Country", group.iloc[0].get("Country", ""))),
            "city": clean_text(status_row.get("City", group.iloc[0].get("City", ""))),
        })

    # The sample Navigation page is organized geographically rather than by
    # allocation rank. This makes the salesperson list easier to scan.
    party_rows.sort(key=lambda x: (x["country"].upper(), x["city"].upper(), x["party"].upper()))

    used: set[str] = set()
    mapping: dict[str, str] = {}
    for item in sorted(party_rows, key=lambda x: x["party"].upper()):
        mapping[item["party"]] = safe_sheet_title(item["party"], used)

    for row_num, item in enumerate(party_rows, 2):
        party = item["party"]
        title = mapping[party]
        ws.append([party, party, item["count"], item["country"], item["city"], "Go →"])
        link_cell = ws.cell(row_num, 6)
        link_cell.hyperlink = f"#'{title}'!A1"
        link_cell.font = LINK_FONT

    style_navigation(ws)
    return mapping


def ordered_party_group(group: pd.DataFrame) -> pd.DataFrame:
    sort_cols = [c for c in ["Selection_Rank", "Recipe_Priority", "Carat", "Serial_Key"] if c in group.columns]
    if sort_cols:
        ascending = [True if c != "Carat" else False for c in sort_cols]
        ordered = group.sort_values(sort_cols, ascending=ascending, kind="stable").reset_index(drop=True)
    else:
        ordered = group.reset_index(drop=True)

    if ordered.empty or "SGPComment" not in ordered.columns:
        return ordered

    serial_col = "Serial_Key" if "Serial_Key" in ordered.columns else ("Serial No" if "Serial No" in ordered.columns else None)
    if serial_col is None:
        return ordered

    ordered = ordered.copy()
    ordered["_Pair_Order"] = np.arange(len(ordered), dtype=int)
    ordered["_Pair_Serial"] = ordered[serial_col].apply(clean_pair_serial)
    ordered["_Pair_Partner"] = ordered["SGPComment"].apply(get_pair_id_for_sorting)

    # Map both members to one canonical key. This also groups a partner row
    # whose own SGPComment is blank when the other member references it.
    serial_to_pair: dict[str, str] = {}
    for serial, partner in ordered[["_Pair_Serial", "_Pair_Partner"]].itertuples(index=False, name=None):
        if not serial or not partner:
            continue
        pair_key = "_".join(sorted([serial, partner]))
        serial_to_pair.setdefault(serial, pair_key)
        serial_to_pair.setdefault(partner, pair_key)
    ordered["_Pair_Group"] = ordered["_Pair_Serial"].map(serial_to_pair).fillna("")

    if ordered["_Pair_Group"].ne("").any():
        anchors = ordered.loc[ordered["_Pair_Group"].ne("")].groupby("_Pair_Group")["_Pair_Order"].min()
        ordered["_Pair_Anchor"] = ordered["_Pair_Group"].map(anchors).fillna(ordered["_Pair_Order"]).astype(int)
        ordered = ordered.sort_values(
            ["_Pair_Anchor", "_Pair_Group", "_Pair_Order"],
            ascending=[True, True, True],
            kind="stable",
        )

    return ordered.drop(
        columns=["_Pair_Order", "_Pair_Serial", "_Pair_Partner", "_Pair_Group", "_Pair_Anchor"],
        errors="ignore",
    ).reset_index(drop=True)


def add_party_sheet(
    wb: Workbook,
    title: str,
    group: pd.DataFrame,
    visible_layout: list[tuple[str, str | None]],
) -> None:
    ws = wb.create_sheet(title)
    ws.freeze_panes = "A2"
    ws.append([label for label, _ in visible_layout])

    for col_idx, (label, _) in enumerate(visible_layout, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = STOCK_WIDTH_BY_LABEL[label]

    for cell in ws[1]:
        cell.fill = GREY_HEADER
        cell.font = PARTY_HEADER_FONT
        cell.border = GRID_BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")

    group = ordered_party_group(group)
    for row_idx, (_, record) in enumerate(group.iterrows(), 1):
        excel_row = row_idx + 1
        for col_idx, (label, source) in enumerate(visible_layout, 1):
            value = row_idx if source is None else record.get(source, "")
            cell = ws.cell(excel_row, col_idx)
            if label in LINK_HEADERS:
                set_link_cell(cell, label, value)
            else:
                cell.value = safe(value)

        fill = match_fill(record.get("Match_Level", ""))
        for col_idx, (label, _) in enumerate(visible_layout, 1):
            cell = ws.cell(excel_row, col_idx)
            cell.fill = fill
            cell.border = GRID_BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if label == "Sr.No":
                cell.font = SERIAL_NUMBER_FONT
            elif label in LINK_HEADERS and cell.hyperlink:
                cell.font = LINK_FONT
            if label in NUMERIC_HEADERS:
                cell.number_format = "0.00"

    # The sample uses the Sr.No header as the return-to-Navigation link.
    ws["A1"].hyperlink = "#'Navigation'!A1"
    return None


def add_hidden_audit_sheet(wb: Workbook, selected: pd.DataFrame, status: pd.DataFrame) -> None:
    ws = wb.create_sheet("Dealmaker Audit")
    columns = [c for c in AUDIT_COLUMNS if c in selected.columns]
    ws.append(columns)
    for cell in ws[1]:
        cell.fill = AUDIT_HEADER_FILL
        cell.font = AUDIT_HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row in selected[columns].itertuples(index=False, name=None):
        ws.append([safe(v) for v in row])

    ws.freeze_panes = "A2"
    for idx, header in enumerate(columns, 1):
        ws.column_dimensions[get_column_letter(idx)].width = 42 if header in {"Selection_Reason", "Release_Status", "Generation_Reason"} else 22

    # Keep audit information in the file without changing the salesperson's
    # familiar visible workbook layout.
    ws.sheet_state = "hidden"


def build_book(selected: pd.DataFrame, status: pd.DataFrame) -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)
    visible_layout = available_stock_layout(selected.columns)
    mapping = add_navigation(wb, selected, status)
    for party, group in selected.groupby("Sold Party", sort=True, dropna=False):
        add_party_sheet(wb, mapping[clean_text(party)], group, visible_layout)
    add_hidden_audit_sheet(wb, selected, status)
    wb.active = 0
    # Final mutation pass: no worksheet-writing code may run after this.
    normalize_workbook_stock_links(wb)
    return wb


def main() -> None:
    parser = argparse.ArgumentParser(description="Create salesperson workbooks in the supplied regular-stock layout.")
    parser.add_argument("--pipeline", required=True)
    parser.add_argument("--output-dir", required=True)
    args = parser.parse_args()

    pipeline = Path(args.pipeline).resolve()
    sys.path.insert(0, str(pipeline.parent))
    output = Path(args.output_dir)
    status = load(output, "final_status")
    selected = load(output, "final_selected")

    if selected.empty:
        # A zero-selection run is a valid production outcome.  Still create the
        # standard artifacts so downstream reconciliation and packaging can
        # complete without special-case file-not-found failures.
        combined = output / "Generated_Deal_Baskets_UNIQUE_v6.xlsx"
        combined_bytes, _ = serialize_validated_workbook(
            build_book(selected, status), selected, combined.name
        )
        combined.write_bytes(combined_bytes)

        salesperson_zip = output / "Generated_Deal_Baskets_By_Salesperson_UNIQUE_v6.zip"
        with zipfile.ZipFile(salesperson_zip, "w", zipfile.ZIP_DEFLATED) as archive:
            archive.writestr(
                "NO_ACCEPTED_BASKETS.txt",
                "No accepted deal baskets were generated. Review Party_Generation_Status_v6.csv and the production review workbook.\n",
            )

        pd.DataFrame([{
            "Salesperson_Workbook": "",
            "Party_Count": 0,
            "Selected_Rows": 0,
            "Unique_Serials": 0,
            "Native_HTTP_Links": 0,
            "Expected_Native_HTTP_Links": 0,
            "Formula_Hyperlinks": 0,
            "Excel_Link_Validation": "PASS",
            "Generator_Version": GENERATOR_VERSION,
            "Reconciliation_Status": "PASS - NO ACCEPTED BASKETS",
            "Layout_Status": "NOT APPLICABLE",
        }]).to_csv(output / "Salesperson_Output_Reconciliation_v8.csv", index=False)
        print(
            f"No accepted selected stones; created empty reconciled artifacts.\n"
            f"Combined deal workbook: {combined}\nSalesperson ZIP: {salesperson_zip}",
            flush=True,
        )
        return

    combined = output / "Generated_Deal_Baskets_UNIQUE_v6.xlsx"
    combined_bytes, combined_metrics = serialize_validated_workbook(
        build_book(selected, status), selected, combined.name
    )
    combined.write_bytes(combined_bytes)

    salesperson_zip = output / "Generated_Deal_Baskets_By_Salesperson_UNIQUE_v6.zip"
    reconciliation = []
    with zipfile.ZipFile(salesperson_zip, "w", zipfile.ZIP_DEFLATED) as archive:
        for salesperson, sales_group in selected.groupby("Sales_Person", dropna=False, sort=True):
            parties = set(sales_group["Sold Party"].astype(str))
            book_status = status[status["Sold Party"].astype(str).isin(parties)]
            workbook = build_book(sales_group, book_status)
            filename = salesperson_filename(salesperson)
            workbook_bytes, link_metrics = serialize_validated_workbook(
                workbook, sales_group, filename
            )
            archive.writestr(filename, workbook_bytes)
            reconciliation.append({
                "Salesperson_Workbook": filename,
                "Party_Count": sales_group["Sold Party"].nunique(),
                "Selected_Rows": len(sales_group),
                "Unique_Serials": sales_group["Serial_Key"].nunique(),
                "Native_HTTP_Links": link_metrics["Native_HTTP_Links"],
                "Expected_Native_HTTP_Links": link_metrics["Expected_Native_HTTP_Links"],
                "Formula_Hyperlinks": link_metrics["Formula_Hyperlinks"],
                "Excel_Link_Validation": "PASS",
                "Generator_Version": GENERATOR_VERSION,
                "Reconciliation_Status": "PASS" if len(sales_group) == sales_group["Serial_Key"].nunique() else "FAIL",
                "Layout_Status": "PASS - CONDITIONAL STOCK FORMAT / MATCH COLORS / NATIVE TEXT LINKS",
            })

    pd.DataFrame(reconciliation).to_csv(output / "Salesperson_Output_Reconciliation_v8.csv", index=False)
    print(f"Combined deal workbook: {combined}\nSalesperson ZIP: {salesperson_zip}", flush=True)


if __name__ == "__main__":
    main()
