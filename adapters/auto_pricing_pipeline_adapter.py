from __future__ import annotations

import io
import json
import os
import subprocess
import sys
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from openpyxl import Workbook, load_workbook

from adapters.common import ROOT_DIR, repository_data_path, safe_filename


FINALIZER_DIR = ROOT_DIR / "processors" / "auto_pricing_finalizer_v4_1_3"
PROCESSOR = FINALIZER_DIR / "process_datasets_v4_1_3.py"
BUILDER = FINALIZER_DIR / "build_auto_pricing_v4_1_3.py"
ALIASES = FINALIZER_DIR / "column_mapping_aliases.json"
BASE_FILES = (
    "Cart.xlsx",
    "Hold.xlsx",
    "Offer.xlsx",
    "View.xlsx",
    "Wishlist.xlsx",
    "sales_120_file.xlsx",
    "sales_180_file.xlsx",
    "Stock file.xlsx",
)

RAW_UPLOAD_NAMES = {
    "Cart": "ADD_TO_CART.xlsx",
    "Hold": "JANGAD_HISTORY.xlsx",
    "Offer": "OFFER_PLACE.xlsx",
    "View": "APPOINTMENT.xlsx",
    "Wishlist": "WISHLIST.xlsx",
    "Sales": "SALE_DATA.xlsx",
    "Stock": "KSTAKA.xlsx",
}

HISTORY_UPLOAD_NAMES = {
    "Cart": "Cart.xlsx",
    "Hold": "Hold.xlsx",
    "Offer": "Offer.xlsx",
    "View": "View.xlsx",
    "Wishlist": "Wishlist.xlsx",
    "Sales 120": "sales_120_file.xlsx",
    "Sales 180": "sales_180_file.xlsx",
}

REQUIRED_RAW_UPLOADS = {"Stock", "Sales"}
REQUIRED_HISTORY_UPLOADS = {"Sales 120", "Sales 180"}
OPTIONAL_ACTION_UPLOADS = {"Cart", "Hold", "Offer", "View", "Wishlist"}



class AutoPricingDataError(RuntimeError):
    """User-facing data preparation error with an optional processing-check workbook."""

    def __init__(self, message: str, *, audit_bytes: bytes | None = None, report: dict | None = None):
        super().__init__(message)
        self.audit_bytes = audit_bytes
        self.report = report or {}


def _sheet_to_single_workbook_bytes(payload: bytes, sheet_name: str) -> bytes:
    """Copy one worksheet by values into a simple standalone workbook."""
    src = load_workbook(io.BytesIO(payload), read_only=True, data_only=False)
    if sheet_name not in src.sheetnames:
        raise ValueError(f"Sheet {sheet_name!r} was not found.")
    src_ws = src[sheet_name]
    out = io.BytesIO()
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Sheet1")
    for row in src_ws.iter_rows(values_only=True):
        ws.append(list(row))
    wb.save(out)
    return out.getvalue()


def _first_sheet_headers(payload: bytes) -> set[str]:
    wb = load_workbook(io.BytesIO(payload), read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
    return {str(v).strip().casefold() for v in row if v not in (None, "")}


def _prepare_previous_sales_file(kind: str, payload: bytes | None) -> bytes | None:
    """Accept either a dedicated Sales base file or the previous combined Auto Pricing workbook."""
    if not payload:
        return payload
    try:
        headers = _first_sheet_headers(payload)
        # Dedicated Sales base files already expose these on the first sheet.
        if "serial no" in headers and "sale date" in headers:
            return payload
        wb = load_workbook(io.BytesIO(payload), read_only=True, data_only=True)
        wanted = "Avg_Rate" if kind == "Sales 120" else "Sale_segment"
        if wanted in wb.sheetnames:
            return _sheet_to_single_workbook_bytes(payload, wanted)
    except Exception as exc:
        raise ValueError(f"The selected {kind} file could not be read. Please choose the matching previous Sales file.") from exc
    raise ValueError(
        f"The selected {kind} file does not contain the expected Sales data. "
        f"Choose the matching previous Sales file, or a previous Auto Pricing workbook containing "
        f"{'Avg_Rate' if kind == 'Sales 120' else 'Sale_segment'}."
    )

ACTION_HEADERS = {
    "Cart": ["CREATED_DATETIME", "Sold Party", "Party_type", "Country", "City", "SALE_PERSON", "Serial No", "Shp", "Carat", "Col", "Qua", "Flo", "Lab", "Cut", "Pol", "Sym", "Shade", "Milky", "TBINC_NAME", "TOIN_NAME", "SBINC_NAME", "SOIN_NAME", "GIRDLE_OPEN_NAME", "PAV_OPEN_NAME", "Ratio", "Disc%", "RAP", "RATE", "AMT"],
    "Offer": ["CREATED_DATETIME", "Sold Party", "Party_type", "Country", "City", "SALE_PERSON", "Serial No", "Shp", "Carat", "Col", "Qua", "Flo", "Lab", "Cut", "Pol", "Sym", "Shade", "Milky", "TBINC_NAME", "TOIN_NAME", "SBINC_NAME", "SOIN_NAME", "GIRDLE_OPEN_NAME", "PAV_OPEN_NAME", "Ratio", "Disc%", "RAP", "RATE", "AMT"],
    "Wishlist": ["CREATED_DATETIME", "Sold Party", "Party_type", "Country", "City", "SALE_PERSON", "Serial No", "Shp", "Carat", "Col", "Qua", "Flo", "Lab", "Cut", "Pol", "Sym", "Shade", "Milky", "TBINC_NAME", "TOIN_NAME", "SBINC_NAME", "SOIN_NAME", "GIRDLE_OPEN_NAME", "PAV_OPEN_NAME", "Ratio", "Disc%", "RAP", "RATE", "AMT"],
    "Hold": ["SERIAL_NO", "JANGAD_DATE", "PARTY_CODE", "Sold Party"],
    "View": ["SEQ_NO", "APPOINTMENT_TYPE", "APPOINTMENT_DATE", "APPOINTMENT_TIME", "PARTY_CODE", "Sold Party", "EMPLOYEE_NO", "SERIAL_NO", "SOURCE", "STATUS", "CREATED_BY", "CREATED_DATETIME", "CREATED_COMPUTER", "MODIFIED_BY", "MODIFIED_DATETIME", "MODIFIED_COMPUTER", "IS_REACHED", "IS_CANCELED", "NOTE", "ASSIGN_TO", "STOCK_REQUEST_TIME", "APPOINTMENT_END_TIME", "IS_DELETE", "IS_LAYOUT", "CABIN_ASSIGN_DATE"],
}


def preview_auto_pricing_column_mapping(
    raw_files: dict[str, bytes | None],
    previous_files: dict[str, bytes | None] | None = None,
    column_overrides: dict | None = None,
) -> list[dict]:
    """Preview deterministic mappings without running the finalizer or altering uploads.

    Missing fields remain explicit. Data-type/date checks are performed by the processor
    during the full run; this preview only validates source header compatibility.
    """
    from processors.auto_pricing_finalizer_v4_1_3.process_datasets_v4_1_3 import (
        STANDARD_SALES_HEADERS, STOCK_FINAL_HEADERS, read_final_headers,
        schema_mapping, resolve_standard_mapping,
    )
    previous_files = previous_files or {}
    output = []
    date_targets = {
        'Cart': 'CREATED_DATETIME', 'Hold': 'JANGAD_DATE',
        'Offer': 'CREATED_DATETIME', 'View': 'APPOINTMENT_DATE',
        'Wishlist': 'CREATED_DATETIME', 'Sales': 'Sale Date',
    }
    for dataset in ('Stock', 'Sales', 'Cart', 'Hold', 'Offer', 'View', 'Wishlist'):
        payload = raw_files.get(dataset)
        if not payload:
            continue
        if dataset == 'Stock':
            targets = STOCK_FINAL_HEADERS
        elif dataset == 'Sales':
            targets = STANDARD_SALES_HEADERS
        elif previous_files.get(dataset):
            book = load_workbook(io.BytesIO(previous_files[dataset]), read_only=True, data_only=False)
            try:
                targets = [str(v).strip() for v in next(book.active.iter_rows(max_row=1, values_only=True))
                           if v is not None and str(v).strip()
                           and ''.join(ch for ch in str(v).lower() if ch.isalnum()) != 'date']
            finally:
                book.close()
        else:
            targets = ACTION_HEADERS[dataset]
        # Preflight only needs the header row. Read every source directly from
        # uploaded bytes: no /mnt/data path or scratch workbook must be writable.
        book = load_workbook(io.BytesIO(payload), read_only=True, data_only=True)
        try:
            raw_header = next(book.active.iter_rows(max_row=1, values_only=True), ())
            source_headers = [str(v).strip() if v is not None else '' for v in raw_header]
        finally:
            book.close()
        if dataset == 'Sales':
            _, rows = resolve_standard_mapping('Sales', targets, source_headers, overrides=column_overrides)
        else:
            rows = schema_mapping(dataset, targets, source_headers, overrides=column_overrides)
        for row in rows:
            target = row['Final Column']
            status = row['Status']
            output.append({
                'Dataset': dataset,
                'Role': 'Business date (window)' if date_targets.get(dataset) == target else 'Field',
                'Final column': target,
                'Source column': row.get('Raw Columns Found') or '—',
                'Status': status,
                'Pricing impact': row.get('Pricing Impact', 'REQUIRED' if status == 'MISSING' else 'OPTIONAL'),
                'Allowed aliases': row.get('Candidate Raw Columns') or '—',
                'Rule': row.get('Rule', ''),
                'Suggestion (review only)': row.get('Suggestion', ''),
                'Method': row.get('Mapping Method', ''),
            })
    return output


def default_remove_party_path() -> Path:
    return repository_data_path("auto_pricing_remove_parties")


def default_remove_party_bytes() -> bytes | None:
    path = default_remove_party_path()
    return path.read_bytes() if path.is_file() else None


def _empty_xlsx_bytes(headers: list[str]) -> bytes:
    payload = io.BytesIO()
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Sheet1")
    ws.append(headers)
    wb.save(payload)
    return payload.getvalue()


def _named_xlsx_zip(
    files: dict[str, bytes | None],
    canonical_names: dict[str, str],
    label: str,
    *,
    required_keys: set[str] | None = None,
    optional_headers: dict[str, list[str]] | None = None,
) -> bytes:
    """Build the workbook bundle expected by the processor from individual uploads."""
    required = set(canonical_names) if required_keys is None else set(required_keys)
    missing = [key for key in required if not files.get(key)]
    if missing:
        raise ValueError(f"{label} is missing required file(s): {', '.join(sorted(missing))}")

    payload = io.BytesIO()
    with zipfile.ZipFile(payload, "w", zipfile.ZIP_DEFLATED) as zf:
        for key, filename in canonical_names.items():
            data = files.get(key)
            if not data:
                headers = (optional_headers or {}).get(key)
                if headers is None:
                    continue
                data = _empty_xlsx_bytes(headers)
            zf.writestr(filename, data)
    return payload.getvalue()


def _extract_xlsx_zip(payload: bytes, target: Path, label: str) -> None:
    """Safely flatten XLSX files from an uploaded ZIP into a temporary directory."""
    target.mkdir(parents=True, exist_ok=True)
    try:
        with zipfile.ZipFile(io.BytesIO(payload), "r") as zf:
            xlsx_members = [
                info for info in zf.infolist()
                if not info.is_dir()
                and info.filename.lower().endswith(".xlsx")
                and not Path(info.filename).name.startswith("~$")
            ]
            if not xlsx_members:
                raise ValueError(f"{label} does not contain any .xlsx workbooks.")
            seen: set[str] = set()
            for info in xlsx_members:
                name = safe_filename(Path(info.filename).name, fallback="input.xlsx")
                key = name.casefold()
                if key in seen:
                    raise ValueError(f"{label} contains duplicate workbook name: {name}")
                seen.add(key)
                (target / name).write_bytes(zf.read(info))
    except zipfile.BadZipFile as exc:
        raise ValueError(f"{label} is not a valid ZIP file.") from exc


def _run(cmd: list[str], *, env: dict[str, str] | None = None) -> str:
    completed = subprocess.run(
        cmd,
        cwd=str(FINALIZER_DIR),
        env=env,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        check=False,
    )
    if completed.returncode != 0:
        raise RuntimeError(
            "Auto Pricing could not be completed because one of the selected files "
            "could not be processed safely. Please review the file status shown above and try again."
        )
    return completed.stdout or ""


def run_auto_pricing_full_pipeline_from_files(
    *,
    raw_files: dict[str, bytes | None],
    previous_base_files: dict[str, bytes | None],
    remove_party_bytes: bytes | None = None,
    test_date: str | None = None,
    column_overrides: dict | None = None,
    party_code_master_bytes: bytes | None = None,
) -> tuple[bytes, str, str]:
    """Run the full pipeline from individual XLSX uploads used by the Streamlit UI."""
    raw_zip = _named_xlsx_zip(
        raw_files,
        RAW_UPLOAD_NAMES,
        "Latest files",
        required_keys=REQUIRED_RAW_UPLOADS,
        optional_headers=ACTION_HEADERS,
    )
    prepared_previous = dict(previous_base_files)
    prepared_previous["Sales 120"] = _prepare_previous_sales_file("Sales 120", prepared_previous.get("Sales 120"))
    prepared_previous["Sales 180"] = _prepare_previous_sales_file("Sales 180", prepared_previous.get("Sales 180"))
    previous_zip = _named_xlsx_zip(
        prepared_previous,
        HISTORY_UPLOAD_NAMES,
        "Previous files",
        required_keys=REQUIRED_HISTORY_UPLOADS,
        optional_headers=ACTION_HEADERS,
    )
    return run_auto_pricing_full_pipeline(
        raw_zip_bytes=raw_zip,
        previous_base_zip_bytes=previous_zip,
        remove_party_bytes=remove_party_bytes,
        test_date=test_date,
        column_overrides=column_overrides,
        party_code_master_bytes=party_code_master_bytes,
    )


def run_auto_pricing_full_pipeline(
    *,
    raw_zip_bytes: bytes,
    previous_base_zip_bytes: bytes,
    remove_party_bytes: bytes | None = None,
    sales_120_history_bytes: bytes | None = None,
    sales_180_history_bytes: bytes | None = None,
    test_date: str | None = None,
    column_overrides: dict | None = None,
    party_code_master_bytes: bytes | None = None,
) -> tuple[bytes, str, str]:
    """Prepare the latest files and build the Auto Pricing workbook."""
    if not PROCESSOR.is_file() or not BUILDER.is_file() or not ALIASES.is_file():
        raise FileNotFoundError("Required Auto Pricing files are unavailable. Please contact the administrator.")

    with tempfile.TemporaryDirectory(prefix="matching_mapper_auto_pricing_full_") as tmp:
        work = Path(tmp)
        raw_dir = work / "raw"
        history_dir = work / "history"
        output_dir = work / "output"
        output_dir.mkdir(parents=True, exist_ok=True)

        _extract_xlsx_zip(raw_zip_bytes, raw_dir, "Current files")
        _extract_xlsx_zip(previous_base_zip_bytes, history_dir, "Previous files")

        if sales_120_history_bytes is not None:
            (history_dir / "sales_120_file.xlsx").write_bytes(sales_120_history_bytes)
        if sales_180_history_bytes is not None:
            (history_dir / "sales_180_file.xlsx").write_bytes(sales_180_history_bytes)


        if remove_party_bytes is not None:
            remove_party_path = work / "remove_parties_override.xlsx"
            remove_party_path.write_bytes(remove_party_bytes)
            removal_source = "uploaded override"
        else:
            remove_party_path = default_remove_party_path()
            if not remove_party_path.is_file():
                raise FileNotFoundError(
                    "The saved Removed Parties list is unavailable. Please contact the administrator."
                )
            removal_source = "saved Removed Parties list"

        env = os.environ.copy()
        env.update({
            "RAW_DATA_DIR": str(raw_dir),
            "FINAL_HISTORY_DIR": str(history_dir),
            "OUTPUT_DIR": str(output_dir),
            "COLUMN_MAPPING_FILE": str(ALIASES),
            "REMOVE_PARTY_FILE": str(remove_party_path),
        })
        # Persisted only inside the isolated run; users can download/re-upload approved rules.
        if column_overrides:
            override_path = work / "approved_column_mapping.json"
            override_path.write_text(json.dumps(column_overrides, indent=2), encoding="utf-8")
            env["COLUMN_MAPPING_OVERRIDE_FILE"] = str(override_path)
        else:
            env.pop("COLUMN_MAPPING_OVERRIDE_FILE", None)
        if party_code_master_bytes:
            master_path = work / "party_code_master.xlsx"
            master_path.write_bytes(party_code_master_bytes)
            env["PARTY_CODE_MASTER_FILE"] = str(master_path)
        else:
            env.pop("PARTY_CODE_MASTER_FILE", None)
        if test_date:
            env["TEST_PROCESSING_DATE"] = test_date
        else:
            env.pop("TEST_PROCESSING_DATE", None)

        logs = []
        logs.append(_run([sys.executable, str(PROCESSOR)], env=env))

        report_path = output_dir / "processing_report.json"
        if not report_path.is_file():
            raise RuntimeError("Processor completed without creating processing_report.json.")
        report = json.loads(report_path.read_text(encoding="utf-8"))
        if not report.get("all_datasets_ok"):
            failed = [
                str(row.get("Dataset"))
                for row in (report.get("summary") or [])
                if str(row.get("Status", "")).upper() != "PASS"
            ]
            errors = [
                row for row in (report.get("issues") or [])
                if str(row.get("Severity", "")).upper() == "ERROR"
            ]
            parts = []
            if failed:
                parts.append("Could not prepare: " + ", ".join(dict.fromkeys(failed)))
            for row in errors[:3]:
                ds = str(row.get("Dataset") or "File")
                issue = str(row.get("Issue Type") or row.get("Details") or "needs attention")
                col = str(row.get("Column") or "").strip()
                parts.append(f"{ds}: {issue}" + (f" ({col})" if col else ""))
            message = ". ".join(parts) or "One or more selected files need attention."
            audit_path = output_dir / "Processing_Audit.xlsx"
            raise AutoPricingDataError(
                message,
                audit_bytes=audit_path.read_bytes() if audit_path.is_file() else None,
                report=report,
            )

        processing_date = str(report.get("processing_date") or "current")
        final_workbook = output_dir / f"Auto_Pricing_{processing_date}.xlsx"
        base_zip = output_dir / f"Updated_Files_{processing_date}.zip"
        logs.append(_run([
            sys.executable,
            str(BUILDER),
            "--base-dir", str(output_dir),
            "--output", str(final_workbook),
            "--base-zip", str(base_zip),
        ]))

        required = [final_workbook, base_zip, output_dir / "Processing_Audit.xlsx", report_path]
        missing = [p.name for p in required if not p.is_file()]
        if missing:
            raise RuntimeError("Processing completed but an expected output is missing: " + ", ".join(missing))

        package_name = f"Auto_Pricing_Output_{processing_date}.zip"
        package_path = work / package_name
        with zipfile.ZipFile(package_path, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.write(final_workbook, arcname=final_workbook.name)
            zf.write(base_zip, arcname=base_zip.name)
            zf.write(output_dir / "Processing_Audit.xlsx", arcname="Processing_Check.xlsx")
            zf.write(report_path, arcname="Run_Summary.json")
            if column_overrides:
                zf.writestr("approved_column_mapping.json", json.dumps(column_overrides, indent=2))
            for filename in BASE_FILES:
                path = output_dir / filename
                if path.is_file():
                    zf.write(path, arcname=f"base_files/{filename}")

        return package_path.read_bytes(), package_name, "\n".join(logs)
