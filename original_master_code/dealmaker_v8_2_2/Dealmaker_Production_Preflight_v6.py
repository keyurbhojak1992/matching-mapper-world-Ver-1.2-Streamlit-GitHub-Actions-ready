#!/usr/bin/env python3
"""Create a production data-quality and mapping report before deal generation."""
from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import Dealmaker_Full_Pipeline_v6 as pipeline
import Dealmaker_Step1_Interval_Engine_v10 as interval
import dealmaker_quality_v6 as quality_domain
from dealmaker_quality_v6 import domain_status
from dealmaker_io_v2 import load_manual_mapping, metadata_rows, read_dataset_table, serial_key_series

VERSION = "8.2.2"
SEVERITY_ORDER = {"ERROR": 0, "WARNING": 1, "OK": 2, "INFO": 3}


def text(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_party_key(value: Any) -> str:
    return re.sub(r"[^A-Z0-9]+", "", text(value).upper())


def add_issue(rows: list[dict[str, Any]], severity: str, dataset: str, issue: str, value: Any, recommendation: str) -> None:
    rows.append({
        "Severity": severity,
        "Dataset": dataset,
        "Issue": issue,
        "Value": value,
        "Recommendation": recommendation,
    })


def analyze_dates(df: pd.DataFrame, dataset: str, issues: list[dict[str, Any]], summary: list[dict[str, Any]]) -> pd.Series:
    parsed = interval.normalize_date(df["Date"])
    invalid = int(parsed.isna().sum())
    add_issue(
        issues,
        "ERROR" if invalid else "OK",
        dataset,
        "Invalid or blank dates",
        invalid,
        "Correct or remove invalid date rows before production." if invalid else "No action required.",
    )
    valid = parsed.dropna()
    summary.append({"Dataset": dataset, "Metric": "Earliest valid date", "Value": valid.min() if not valid.empty else ""})
    summary.append({"Dataset": dataset, "Metric": "Latest valid date", "Value": valid.max() if not valid.empty else ""})
    future = int((parsed > (pd.Timestamp.today().normalize() + pd.Timedelta(days=1))).sum())
    add_issue(
        issues,
        "ERROR" if future else "OK",
        dataset,
        "Future-dated rows",
        future,
        "Verify date format and source system dates." if future else "No action required.",
    )
    return parsed


def party_variant_review(frames: dict[str, pd.DataFrame]) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    combined = []
    for dataset, df in frames.items():
        if "Sold Party" not in df.columns:
            continue
        part = df[["Sold Party"]].copy()
        part["Dataset"] = dataset
        part["Sold Party"] = part["Sold Party"].fillna("").astype(str).str.strip()
        part = part[part["Sold Party"].ne("")]
        part["Party_Normalized_Key"] = part["Sold Party"].map(normalize_party_key)
        combined.append(part)
    if not combined:
        return pd.DataFrame(columns=["Party_Normalized_Key", "Variant_Count", "Variants", "Datasets", "Total_Rows"])
    data = pd.concat(combined, ignore_index=True)
    for key, g in data.groupby("Party_Normalized_Key"):
        variants = sorted(g["Sold Party"].unique())
        if len(variants) > 1:
            rows.append({
                "Party_Normalized_Key": key,
                "Variant_Count": len(variants),
                "Variants": " | ".join(variants),
                "Datasets": ", ".join(sorted(g["Dataset"].unique())),
                "Total_Rows": len(g),
            })
    return pd.DataFrame(rows).sort_values(["Variant_Count", "Total_Rows"], ascending=False) if rows else pd.DataFrame(columns=["Party_Normalized_Key", "Variant_Count", "Variants", "Datasets", "Total_Rows"])


def serial_review(frames: dict[str, pd.DataFrame], sale_types: set[str]) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    keyed: dict[str, pd.Series] = {}
    for dataset, df in frames.items():
        if "Serial No" not in df.columns:
            continue
        keys = serial_key_series(df["Serial No"])
        keyed[dataset] = keys
        nonblank = keys[keys.ne("")]
        duplicate_rows = int(nonblank.duplicated(keep=False).sum())
        duplicate_keys = int(nonblank[nonblank.duplicated(keep=False)].nunique())
        rows.append({"Check": "Duplicate serials within dataset", "Dataset_A": dataset, "Dataset_B": "", "Rows_or_Serials": duplicate_rows, "Unique_Serials": duplicate_keys, "Severity": "WARNING" if duplicate_rows else "OK"})
    names = sorted(keyed)
    for i, left in enumerate(names):
        for right in names[i + 1:]:
            overlap = set(keyed[left][keyed[left].ne("")]) & set(keyed[right][keyed[right].ne("")])
            severity = "INFO"
            if {left, right} == {"Sales", "Stock"} and overlap:
                severity = "WARNING"
            rows.append({"Check": "Serial overlap across datasets", "Dataset_A": left, "Dataset_B": right, "Rows_or_Serials": len(overlap), "Unique_Serials": len(overlap), "Severity": severity})
    return pd.DataFrame(rows)


def top_values(df: pd.DataFrame, dataset: str, columns: list[str], max_values: int = 40) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    for col in columns:
        if col not in df.columns:
            continue
        counts = df[col].fillna("<BLANK>").astype(str).str.strip().replace("", "<BLANK>").value_counts(dropna=False).head(max_values)
        for value, count in counts.items():
            rows.append({"Dataset": dataset, "Column": col, "Value": value, "Count": int(count)})
    return pd.DataFrame(rows)


def style_workbook(path: Path) -> None:
    wb = load_workbook(path)
    header_fill = PatternFill("solid", fgColor="1F4E78")
    error_fill = PatternFill("solid", fgColor="F4CCCC")
    warning_fill = PatternFill("solid", fgColor="FCE5CD")
    ok_fill = PatternFill("solid", fgColor="D9EAD3")
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.sheet_view.showGridLines = False
        if ws.max_row >= 1:
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = Font(color="FFFFFF", bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.auto_filter.ref = ws.dimensions
        for col_idx in range(1, ws.max_column + 1):
            max_len = 0
            for row_idx in range(1, min(ws.max_row, 250) + 1):
                value = ws.cell(row_idx, col_idx).value
                max_len = max(max_len, len(str(value)) if value is not None else 0)
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 55)
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
        if ws.title == "Issues" and ws.max_row >= 2:
            headers = {cell.value: cell.column for cell in ws[1]}
            sev_col = headers.get("Severity")
            if sev_col:
                for r in range(2, ws.max_row + 1):
                    sev = str(ws.cell(r, sev_col).value or "").upper()
                    fill = error_fill if sev == "ERROR" else warning_fill if sev == "WARNING" else ok_fill if sev == "OK" else None
                    if fill:
                        for c in range(1, ws.max_column + 1):
                            ws.cell(r, c).fill = fill
    wb.save(path)


def run(args: argparse.Namespace) -> int:
    sales_path = Path(args.sales)
    actions_path = Path(args.actions)
    stock_path = Path(args.stock)
    control_path = Path(args.control)
    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    settings, _, _, _ = pipeline.load_control(control_path)
    quality_config = quality_domain.configure_quality(control_path)
    pipeline.configure_quality(control_path)
    if hasattr(interval, "configure_quality"):
        interval.configure_quality(control_path)
    sale_types = pipeline.parse_csv_set(settings.get("SALE_TYPE_VALUES", "SALE,SALES"))
    bid_types = pipeline.parse_csv_set(settings.get("BID_TYPE_VALUES", "BID"))

    configs = {
        "Sales": (sales_path, interval.COLUMN_ALIASES, ["Serial No", "Sold Party", "Date", "Type"], args.sales_sheet or None, args.sales_header_row, load_manual_mapping(control_path, "SALES")),
        "Actions": (actions_path, interval.COLUMN_ALIASES, ["Serial No", "Sold Party", "Date", "Type"], args.actions_sheet or None, args.actions_header_row, load_manual_mapping(control_path, "ACTIONS")),
        "Stock": (stock_path, pipeline.STOCK_ALIASES, ["Serial No", "Shp", "Carat", "Col", "Qua", "Flo", "Lab"], args.stock_sheet or None, args.stock_header_row, load_manual_mapping(control_path, "STOCK")),
    }

    issues: list[dict[str, Any]] = []
    summary: list[dict[str, Any]] = [
        {"Dataset": "Run", "Metric": "Preflight version", "Value": VERSION},
        {"Dataset": "Run", "Metric": "Control workbook", "Value": control_path.name},
    ]
    frames: dict[str, pd.DataFrame] = {}
    mapping_rows: list[dict[str, Any]] = []
    inventory_rows: list[dict[str, Any]] = []

    for dataset, (path, aliases, required, sheet_name, header_row, mapping) in configs.items():
        try:
            df, meta = read_dataset_table(path, aliases, required, dataset, sheet_name, header_row, mapping)
            frames[dataset] = df
            mapping_rows.extend(metadata_rows(meta))
            for idx, col in enumerate(meta.raw_columns, 1):
                inventory_rows.append({"Dataset": dataset, "Column_Order": idx, "Source_Column": col, "Standardized_Column": meta.resolved_mapping.get(col, col)})
            summary.extend([
                {"Dataset": dataset, "Metric": "Source file", "Value": path.name},
                {"Dataset": dataset, "Metric": "Selected worksheet", "Value": meta.sheet_name or "Delimited file"},
                {"Dataset": dataset, "Metric": "Detected header row (Excel numbering)", "Value": meta.header_row_zero_based + 1},
                {"Dataset": dataset, "Metric": "Rows", "Value": len(df)},
                {"Dataset": dataset, "Metric": "Columns", "Value": len(df.columns)},
            ])
            add_issue(issues, "OK", dataset, "Required columns mapped", ", ".join(required), "No action required.")
        except Exception as exc:
            frames[dataset] = pd.DataFrame()
            add_issue(issues, "ERROR", dataset, "Input could not be mapped/read", str(exc), "Use Column_Mapping or explicit worksheet/header settings, then rerun preflight.")

    # Sales and Actions checks
    for dataset in ["Sales", "Actions"]:
        df = frames.get(dataset, pd.DataFrame())
        if df.empty:
            continue
        df = df.copy()
        df["Date_Parsed"] = analyze_dates(df, dataset, issues, summary)
        df["Party_Clean"] = df["Sold Party"].fillna("").astype(str).str.strip()
        df["Serial_Key"] = serial_key_series(df["Serial No"])
        df["Type_Clean"] = df["Type"].fillna("").astype(str).str.strip().str.upper()
        blank_party = int(df["Party_Clean"].eq("").sum())
        blank_serial = int(df["Serial_Key"].eq("").sum())
        add_issue(issues, "ERROR" if blank_party else "OK", dataset, "Blank party names", blank_party, "Fill party names or remove unusable rows." if blank_party else "No action required.")
        add_issue(issues, "WARNING" if blank_serial else "OK", dataset, "Blank serial numbers", blank_serial, "Blank serial rows cannot support exact serial conversion or unique-stone counts." if blank_serial else "No action required.")
        dup_mask = df.duplicated(["Party_Clean", "Date_Parsed", "Type_Clean", "Serial_Key"], keep=False)
        dup_participating = int(dup_mask.sum())
        dup_excess = int(len(df) - len(df.drop_duplicates(["Party_Clean", "Date_Parsed", "Type_Clean", "Serial_Key"], keep="first")))
        add_issue(issues, "WARNING" if dup_excess else "OK", dataset,
                  "Excess duplicate rows that production will remove", dup_excess,
                  f"{dup_participating} rows participate in duplicate groups. Confirm repeated actions versus repeated exports." if dup_excess else "No action required.")
        frames[dataset] = df

    sales = frames.get("Sales", pd.DataFrame())
    actions = frames.get("Actions", pd.DataFrame())
    stock = frames.get("Stock", pd.DataFrame())
    stock_sale_overlap = pd.DataFrame()
    discount_outlier_parts: list[pd.DataFrame] = []

    if not sales.empty:
        sale_count = int(sales["Type_Clean"].isin(sale_types).sum())
        bid_count = int(sales["Type_Clean"].isin(bid_types).sum())
        unknown_types = sorted(set(sales["Type_Clean"].dropna()) - sale_types - bid_types - {""})
        add_issue(issues, "ERROR" if sale_count == 0 else "OK", "Sales", "Completed SALE/SALES rows", sale_count, "Update SALE_TYPE_VALUES or correct transaction types." if sale_count == 0 else "No action required.")
        add_issue(issues, "INFO", "Sales", "BID rows (intent only)", bid_count, "BID rows will not count as completed purchases.")
        add_issue(issues, "WARNING" if unknown_types else "OK", "Sales", "Other transaction types", ", ".join(unknown_types) or "None", "Review and map only genuine completed-sale values into SALE_TYPE_VALUES.")

    if not actions.empty:
        canonical = pipeline.canonicalize_action_types(actions.rename(columns={"Type_Clean": "Type_Source"}).assign(Type=actions["Type_Clean"]), settings)
        frames["Actions"] = canonical
        action_types = sorted(canonical["Type"].dropna().astype(str).unique())
        unknown_actions = sorted(set(action_types) - set(pipeline.ACTION_TYPES))
        add_issue(issues, "WARNING" if unknown_actions else "OK", "Actions", "Unmapped action types", ", ".join(unknown_actions) or "None", "Add source names to HOLD/CART/OFFER/WISHLIST/DETAILS/VIEW_TYPE_VALUES.")
        for wanted in ["HOLD", "CART", "OFFER", "WISHLIST", "DETAILS", "VIEW"]:
            count = int(canonical["Type"].eq(wanted).sum())
            severity = "INFO"
            add_issue(issues, severity, "Actions", f"{wanted} rows", count, "Behavior can only be learned when this action exists.")

    if not stock.empty:
        stock = stock.copy()
        stock["Serial_Key"] = serial_key_series(stock["Serial No"])
        blank_serial = int(stock["Serial_Key"].eq("").sum())
        add_issue(issues, "ERROR" if blank_serial else "OK", "Stock", "Blank serial numbers", blank_serial, "Every current-stock row must have a serial number." if blank_serial else "No action required.")
        duplicate = int(stock[stock["Serial_Key"].ne("")].duplicated("Serial_Key", keep=False).sum())
        add_issue(issues, "WARNING" if duplicate else "OK", "Stock", "Duplicate current-stock serial rows", duplicate, "Resolve duplicates or confirm which row is current." if duplicate else "No action required.")
        for col in ["Shp", "Carat", "Col", "Qua", "Flo", "Lab"]:
            blank = int(stock[col].fillna("").astype(str).str.strip().eq("").sum())
            add_issue(issues, "ERROR" if blank else "OK", "Stock", f"Blank required specification: {col}", blank, "Fill required matching fields." if blank else "No action required.")
        status_exists = "Stock_Status" in stock.columns and stock["Stock_Status"].fillna("").astype(str).str.strip().ne("").any()
        require_status = pipeline.yes(settings.get("REQUIRE_STOCK_STATUS_COLUMN", "NO"))
        add_issue(
            issues,
            "ERROR" if require_status and not status_exists else "WARNING" if not status_exists else "OK",
            "Stock",
            "Availability/status field",
            "Detected" if status_exists else "Not detected or blank",
            "Map the source status column to Stock_Status and configure available/unavailable values." if not status_exists else "No action required.",
        )
        if status_exists:
            statuses = stock["Stock_Status"].fillna("").astype(str).str.strip().str.upper()
            available = pipeline.parse_csv_set(settings.get("STOCK_AVAILABLE_VALUES", ""))
            unavailable = pipeline.parse_csv_set(settings.get("STOCK_UNAVAILABLE_VALUES", ""))
            unknown = sorted(set(statuses[statuses.ne("")]) - available - unavailable)
            add_issue(issues, "WARNING" if unknown else "OK", "Stock", "Unknown stock status values", ", ".join(unknown) or "None", "Add values to STOCK_AVAILABLE_VALUES or STOCK_UNAVAILABLE_VALUES.")
        frames["Stock"] = stock

    # Cross-file checks
    if not sales.empty and not stock.empty:
        completed = sales[sales["Type_Clean"].isin(sale_types)]
        overlap = set(completed["Serial_Key"]) & set(stock["Serial_Key"])
        overlap.discard("")
        stock_sale_overlap = stock[stock["Serial_Key"].isin(overlap)].copy()
        add_issue(issues, "WARNING" if overlap else "OK", "Cross-file", "Current-stock serials found in completed-sale history", len(overlap), "Review the Stock_Sale_Overlap sheet. The v7 pipeline excludes these serials by default.")

    party_variants = party_variant_review({k: v for k, v in frames.items() if k in {"Sales", "Actions"}})
    add_issue(issues, "WARNING" if len(party_variants) else "OK", "Cross-file", "Possible party-name variants", len(party_variants), "Standardize party names or use a controlled source-system identifier.")

    serials = serial_review({k: v for k, v in frames.items() if k in {"Sales", "Actions", "Stock"}}, sale_types)

    type_frames = []
    for dataset, df in frames.items():
        type_col = "Type" if "Type" in df.columns else None
        if type_col:
            counts = df[type_col].fillna("<BLANK>").astype(str).str.strip().replace("", "<BLANK>").value_counts().rename_axis("Value").reset_index(name="Count")
            counts.insert(0, "Column", "Type")
            counts.insert(0, "Dataset", dataset)
            type_frames.append(counts)
    type_counts = pd.concat(type_frames, ignore_index=True) if type_frames else pd.DataFrame(columns=["Dataset", "Column", "Value", "Count"])

    domain_frames = []
    domain_cols = ["Shade", "Milky", "TOIN_NAME", "SOIN_NAME", "GIRDLE_OPEN_NAME", "PAV_OPEN_NAME", "TBINC_NAME", "SBINC_NAME", "Stock_Status"]
    for dataset, df in frames.items():
        domain_frames.append(top_values(df, dataset, domain_cols))
    value_domains = pd.concat([x for x in domain_frames if not x.empty], ignore_index=True) if any(not x.empty for x in domain_frames) else pd.DataFrame(columns=["Dataset", "Column", "Value", "Count"])

    # Fail-closed quality-domain audit. Every value used by the quality engine
    # must be canonical or explicitly aliased; unmapped values block production.
    quality_domain_rows = []
    quality_columns = {
        "Shade": ["Shade"],
        "Milky": ["Milky"],
        "Open": ["TOIN_NAME", "SOIN_NAME", "GIRDLE_OPEN_NAME", "PAV_OPEN_NAME"],
        "Blk": ["TBINC_NAME", "SBINC_NAME"],
    }
    for dataset, df in frames.items():
        if df.empty:
            continue
        for dimension, columns in quality_columns.items():
            for col in columns:
                if col not in df.columns:
                    continue
                counts = df[col].fillna("").astype(str).value_counts(dropna=False)
                for raw_value, count in counts.items():
                    raw, canonical, status = domain_status(raw_value, dimension)
                    quality_domain_rows.append({
                        "Dataset": dataset, "Dimension": dimension, "Source_Column": col,
                        "Raw_Value": raw or "<BLANK>", "Canonical_Value": canonical or "<BLANK>",
                        "Status": status, "Count": int(count),
                    })
    quality_domain_audit = pd.DataFrame(quality_domain_rows)
    if not quality_domain_audit.empty:
        unmapped = quality_domain_audit[quality_domain_audit["Status"].eq("UNMAPPED")]
        blank = quality_domain_audit[quality_domain_audit["Status"].eq("UNKNOWN")]
        policy = str(settings.get("QUALITY_DOMAIN_POLICY", "ERROR")).strip().upper()
        add_issue(issues, "ERROR" if len(unmapped) and policy == "ERROR" else "WARNING" if len(unmapped) else "OK",
                  "Quality domains", "Unmapped quality values", int(unmapped["Count"].sum()) if len(unmapped) else 0,
                  "Map every value in Quality_Domain_Audit before production. v7 never silently scores an unknown quality value.")
        add_issue(issues, "WARNING" if int(blank["Count"].sum()) else "OK", "Quality domains",
                  "Blank/unknown quality cells across source quality columns", int(blank["Count"].sum()),
                  "The count is cells, not rows. Current-stock rows with unresolved required quality values are excluded; historical unknowns use controlled fallbacks.")
    else:
        quality_domain_audit = pd.DataFrame(columns=["Dataset","Dimension","Source_Column","Raw_Value","Canonical_Value","Status","Count"])

    # Discount scale profile
    for dataset, df in frames.items():
        disc_col = next((c for c in ["Disc%", "Discount", "Dis%"] if c in df.columns), None)
        if disc_col:
            values = pd.to_numeric(df[disc_col], errors="coerce").dropna()
            if not values.empty:
                summary.extend([
                    {"Dataset": dataset, "Metric": "Discount minimum", "Value": float(values.min())},
                    {"Dataset": dataset, "Metric": "Discount median", "Value": float(values.median())},
                    {"Dataset": dataset, "Metric": "Discount maximum", "Value": float(values.max())},
                ])
                outlier_mask = pd.to_numeric(df[disc_col], errors="coerce").abs().gt(100)
                outlier_count = int(outlier_mask.sum())
                add_issue(issues, "WARNING" if outlier_count else "OK", dataset, "Discount values outside -100 to 100", outlier_count, "Review Discount_Outliers; verify whether these are valid discount points or source-data errors.")
                if outlier_count:
                    keep = [c for c in ["Serial No", "Sold Party", "Date", "Type", "Shp", "Carat", "Col", "Qua", "Flo", "Lab", disc_col, "Rate", "Amount", "AMT", "Loc"] if c in df.columns]
                    part = df.loc[outlier_mask, keep].copy()
                    part.insert(0, "Dataset", dataset)
                    discount_outlier_parts.append(part)
                if values.abs().quantile(0.95) <= 1.5:
                    add_issue(issues, "WARNING", dataset, "Discount scale appears fractional", "95% of absolute values <= 1.5", "Confirm whether discounts are stored as fractions (e.g., -0.35) or points (e.g., -35).")

    issue_df = pd.DataFrame(issues)
    if not issue_df.empty:
        issue_df["_order"] = issue_df["Severity"].map(SEVERITY_ORDER).fillna(99)
        issue_df = issue_df.sort_values(["_order", "Dataset", "Issue"]).drop(columns="_order").reset_index(drop=True)
    error_count = int(issue_df["Severity"].eq("ERROR").sum()) if not issue_df.empty else 0
    warning_count = int(issue_df["Severity"].eq("WARNING").sum()) if not issue_df.empty else 0
    summary[0:0] = [
        {"Dataset": "Run", "Metric": "Blocking ERROR checks", "Value": error_count},
        {"Dataset": "Run", "Metric": "WARNING checks", "Value": warning_count},
        {"Dataset": "Run", "Metric": "Preflight status", "Value": "BLOCKED" if error_count else "READY FOR CONTROLLED PRODUCTION RUN"},
    ]

    discount_outliers = pd.concat(discount_outlier_parts, ignore_index=True) if discount_outlier_parts else pd.DataFrame(columns=["Dataset", "Serial No", "Disc%"] )

    previews = {}
    for dataset, df in frames.items():
        previews[f"Preview_{dataset}"] = df.head(20).copy()

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        pd.DataFrame(summary).to_excel(writer, sheet_name="Summary", index=False)
        issue_df.to_excel(writer, sheet_name="Issues", index=False)
        pd.DataFrame(mapping_rows).to_excel(writer, sheet_name="Column_Mapping", index=False)
        pd.DataFrame(inventory_rows).to_excel(writer, sheet_name="Column_Inventory", index=False)
        type_counts.to_excel(writer, sheet_name="Type_Counts", index=False)
        party_variants.to_excel(writer, sheet_name="Party_Name_Review", index=False)
        serials.to_excel(writer, sheet_name="Serial_Review", index=False)
        stock_sale_overlap.to_excel(writer, sheet_name="Stock_Sale_Overlap", index=False)
        discount_outliers.to_excel(writer, sheet_name="Discount_Outliers", index=False)
        value_domains.to_excel(writer, sheet_name="Value_Domains", index=False)
        quality_domain_audit.to_excel(writer, sheet_name="Quality_Domain_Audit", index=False)
        quality_config.domain.to_excel(writer, sheet_name="Quality_Domain_Config", index=False)
        quality_config.aliases.to_excel(writer, sheet_name="Quality_Alias_Config", index=False)
        for sheet, df in previews.items():
            df.to_excel(writer, sheet_name=sheet[:31], index=False)
    style_workbook(output_path)

    print(f"Preflight workbook: {output_path}")
    print(f"Blocking errors: {error_count} | Warnings: {warning_count}")
    if error_count:
        print("Production should not run until ERROR checks are corrected or explicitly reviewed.")
    else:
        print("Preflight has no blocking errors. Review warnings before production generation.")
    return error_count


def parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Inspect and map Dealmaker production inputs before generation.")
    p.add_argument("--sales", required=True)
    p.add_argument("--actions", required=True)
    p.add_argument("--stock", required=True)
    p.add_argument("--control", required=True)
    p.add_argument("--output", default="Dealmaker_Data_Preflight.xlsx")
    p.add_argument("--sales-sheet", default="")
    p.add_argument("--actions-sheet", default="")
    p.add_argument("--stock-sheet", default="")
    p.add_argument("--sales-header-row", type=int, default=None)
    p.add_argument("--actions-header-row", type=int, default=None)
    p.add_argument("--stock-header-row", type=int, default=None)
    return p


if __name__ == "__main__":
    errors = run(parser().parse_args())
    # The workbook is always produced. Exit 0 so Colab can display it; the next
    # cell reads the Issues sheet and blocks production when ERROR rows exist.
    raise SystemExit(0)
