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
    "Stock": "Stock file.xlsx",
}


def default_remove_party_path() -> Path:
    return repository_data_path("auto_pricing_remove_parties")


def default_remove_party_bytes() -> bytes | None:
    path = default_remove_party_path()
    return path.read_bytes() if path.is_file() else None


def default_auto_pricing_template_path() -> Path:
    return repository_data_path("auto_pricing_template")


def default_auto_pricing_template_bytes() -> bytes | None:
    path = default_auto_pricing_template_path()
    return path.read_bytes() if path.is_file() else None


def _named_xlsx_zip(
    files: dict[str, bytes | None],
    canonical_names: dict[str, str],
    label: str,
) -> bytes:
    """Build the same ZIP structure expected by the finalizer from individual uploads."""
    missing = [key for key in canonical_names if not files.get(key)]
    if missing:
        raise ValueError(f"{label} is missing required file(s): {', '.join(missing)}")

    payload = io.BytesIO()
    with zipfile.ZipFile(payload, "w", zipfile.ZIP_DEFLATED) as zf:
        for key, filename in canonical_names.items():
            zf.writestr(filename, files[key] or b"")
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
        tail = (completed.stdout or "")[-8000:]
        raise RuntimeError(f"Auto Pricing pipeline failed.\n{tail}")
    return completed.stdout or ""


def run_auto_pricing_full_pipeline_from_files(
    *,
    raw_files: dict[str, bytes | None],
    previous_base_files: dict[str, bytes | None],
    remove_party_bytes: bytes | None = None,
    test_date: str | None = None,
) -> tuple[bytes, str, str]:
    """Run the full pipeline from individual XLSX uploads used by the Streamlit UI."""
    raw_zip = _named_xlsx_zip(raw_files, RAW_UPLOAD_NAMES, "Latest RAW uploads")
    previous_zip = _named_xlsx_zip(
        previous_base_files, HISTORY_UPLOAD_NAMES, "Previous finalized base uploads"
    )
    return run_auto_pricing_full_pipeline(
        raw_zip_bytes=raw_zip,
        previous_base_zip_bytes=previous_zip,
        template_bytes=None,
        remove_party_bytes=remove_party_bytes,
        test_date=test_date,
    )


def run_auto_pricing_full_pipeline(
    *,
    raw_zip_bytes: bytes,
    previous_base_zip_bytes: bytes,
    template_bytes: bytes | None = None,
    remove_party_bytes: bytes | None = None,
    sales_120_history_bytes: bytes | None = None,
    sales_180_history_bytes: bytes | None = None,
    test_date: str | None = None,
) -> tuple[bytes, str, str]:
    """Run Dataset Finalizer v4.1.3 and then build the formatted Auto Pricing workbook."""
    if not PROCESSOR.is_file() or not BUILDER.is_file() or not ALIASES.is_file():
        raise FileNotFoundError("Bundled Auto Pricing v4.1.3 processor files are missing.")

    with tempfile.TemporaryDirectory(prefix="matching_mapper_auto_pricing_full_") as tmp:
        work = Path(tmp)
        raw_dir = work / "raw"
        history_dir = work / "history"
        output_dir = work / "output"
        output_dir.mkdir(parents=True, exist_ok=True)

        _extract_xlsx_zip(raw_zip_bytes, raw_dir, "RAW ZIP")
        _extract_xlsx_zip(previous_base_zip_bytes, history_dir, "Previous base ZIP")

        if sales_120_history_bytes is not None:
            (history_dir / "sales_120_file.xlsx").write_bytes(sales_120_history_bytes)
        if sales_180_history_bytes is not None:
            (history_dir / "sales_180_file.xlsx").write_bytes(sales_180_history_bytes)

        template_path = work / "Auto_Pricing_Template.xlsx"
        if template_bytes is None:
            bundled_template = default_auto_pricing_template_path()
            if not bundled_template.is_file():
                raise FileNotFoundError(
                    "The Git default Auto Pricing template is missing: "
                    "data/auto_pricing_template.xlsx"
                )
            template_path.write_bytes(bundled_template.read_bytes())
            template_source = "Git default data/auto_pricing_template.xlsx"
        else:
            template_path.write_bytes(template_bytes)
            template_source = "uploaded template override"

        if remove_party_bytes is not None:
            remove_party_path = work / "remove_parties_override.xlsx"
            remove_party_path.write_bytes(remove_party_bytes)
            removal_source = "uploaded override"
        else:
            remove_party_path = default_remove_party_path()
            if not remove_party_path.is_file():
                raise FileNotFoundError(
                    "No Removed Parties override was uploaded and the Git default "
                    "data/remove_parties.xlsx is missing."
                )
            removal_source = "Git default data/remove_parties.xlsx"

        env = os.environ.copy()
        env.update({
            "RAW_DATA_DIR": str(raw_dir),
            "FINAL_HISTORY_DIR": str(history_dir),
            "OUTPUT_DIR": str(output_dir),
            "COLUMN_MAPPING_FILE": str(ALIASES),
            "REMOVE_PARTY_FILE": str(remove_party_path),
        })
        if test_date:
            env["TEST_PROCESSING_DATE"] = test_date
        else:
            env.pop("TEST_PROCESSING_DATE", None)

        logs = []
        logs.append(f"Auto Pricing template source: {template_source}")
        logs.append(f"Removed Parties source: {removal_source}")
        logs.append(_run([sys.executable, str(PROCESSOR)], env=env))

        report_path = output_dir / "processing_report.json"
        if not report_path.is_file():
            raise RuntimeError("Processor completed without creating processing_report.json.")
        report = json.loads(report_path.read_text(encoding="utf-8"))
        if not report.get("all_datasets_ok"):
            raise RuntimeError(
                "One or more base datasets failed validation. Review Processing_Audit.xlsx."
            )

        processing_date = str(report.get("processing_date") or "current")
        final_workbook = output_dir / f"Auto_pricing_FINAL_{processing_date}_v4_1_3.xlsx"
        base_zip = output_dir / f"Final_Base_Output_{processing_date}_v4_1_3.zip"
        logs.append(_run([
            sys.executable,
            str(BUILDER),
            "--template", str(template_path),
            "--base-dir", str(output_dir),
            "--output", str(final_workbook),
            "--base-zip", str(base_zip),
        ]))

        required = [final_workbook, base_zip, output_dir / "Processing_Audit.xlsx", report_path]
        missing = [p.name for p in required if not p.is_file()]
        if missing:
            raise RuntimeError("Pipeline completed but expected output is missing: " + ", ".join(missing))

        package_name = f"Auto_Pricing_Full_Output_{processing_date}_v4_1_3.zip"
        package_path = work / package_name
        with zipfile.ZipFile(package_path, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.write(final_workbook, arcname=final_workbook.name)
            zf.write(base_zip, arcname=base_zip.name)
            zf.write(output_dir / "Processing_Audit.xlsx", arcname="Processing_Audit.xlsx")
            zf.write(report_path, arcname="processing_report.json")
            for filename in BASE_FILES:
                path = output_dir / filename
                if path.is_file():
                    zf.write(path, arcname=f"base_files/{filename}")

        return package_path.read_bytes(), package_name, "\n".join(logs)
