from __future__ import annotations

import io
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path, PurePosixPath

from adapters.common import MASTER_CODE_DIR


PACKAGE_DIR = MASTER_CODE_DIR / "autoprice_v4_1_3"
RUNNER = PACKAGE_DIR / "run_all_v4_1_3.py"
AUTOPRICE_VERSION = "4.1.3"


def _validate_zip_bytes(payload: bytes, label: str) -> None:
    """Reject malformed/path-traversal ZIPs before the supplied runner extracts them."""
    try:
        with zipfile.ZipFile(io.BytesIO(payload)) as zf:
            members = zf.infolist()
            if not members:
                raise ValueError(f"{label} is empty.")
            for info in members:
                path = PurePosixPath(info.filename.replace("\\", "/"))
                if path.is_absolute() or ".." in path.parts:
                    raise ValueError(f"{label} contains an unsafe path: {info.filename}")
    except zipfile.BadZipFile as exc:
        raise ValueError(f"{label} is not a valid ZIP file.") from exc


def _write_optional(path: Path, payload: bytes | None) -> str | None:
    if payload is None:
        return None
    path.write_bytes(payload)
    return str(path)


def run_auto_pricing(
    *,
    raw_zip_bytes: bytes,
    previous_final_zip_bytes: bytes,
    template_bytes: bytes,
    remove_party_bytes: bytes | None = None,
    sales_120_history_bytes: bytes | None = None,
    sales_180_history_bytes: bytes | None = None,
) -> tuple[bytes, str, str]:
    """Run the supplied Dataset Finalizer + Auto Pricing v4.1.3 end-to-end flow."""
    _validate_zip_bytes(raw_zip_bytes, "Raw ZIP")
    _validate_zip_bytes(previous_final_zip_bytes, "Previous finalized-base ZIP")

    if not RUNNER.is_file():
        raise RuntimeError(f"Auto Pricing v{AUTOPRICE_VERSION} runner is missing: {RUNNER}")

    with tempfile.TemporaryDirectory(prefix="matching_mapper_autoprice_v4_1_3_") as tmp:
        work = Path(tmp)
        raw_zip = work / "Raw.zip"
        input_zip = work / "Input latest.zip"
        template = work / "Latest Auto pricing.xlsx"
        output_dir = work / "output"

        raw_zip.write_bytes(raw_zip_bytes)
        input_zip.write_bytes(previous_final_zip_bytes)
        template.write_bytes(template_bytes)

        remove_party = _write_optional(work / "Remove party.xlsx", remove_party_bytes)
        sales_120 = _write_optional(work / "sales_120_file_input.xlsx", sales_120_history_bytes)
        sales_180 = _write_optional(work / "sales_180_file_input.xlsx", sales_180_history_bytes)

        command = [
            sys.executable,
            str(RUNNER),
            "--raw-zip",
            str(raw_zip),
            "--input-zip",
            str(input_zip),
            "--template",
            str(template),
            "--output-dir",
            str(output_dir),
        ]
        if remove_party:
            command.extend(["--remove-party", remove_party])
        if sales_120:
            command.extend(["--sales-120-history", sales_120])
        if sales_180:
            command.extend(["--sales-180-history", sales_180])

        completed = subprocess.run(
            command,
            cwd=str(PACKAGE_DIR),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            check=False,
        )
        logs = completed.stdout or ""
        if completed.returncode != 0:
            raise RuntimeError(
                f"Auto Pricing v{AUTOPRICE_VERSION} failed with exit code {completed.returncode}.\n"
                + logs[-12000:]
            )

        final_books = sorted(output_dir.glob("Auto_pricing_FINAL_*_v4_1_3.xlsx"))
        base_zips = sorted(output_dir.glob("Final_Base_Output_*_v4_1_3.zip"))
        audit = output_dir / "Processing_Audit.xlsx"
        report = output_dir / "processing_report.json"

        missing = []
        if len(final_books) != 1:
            missing.append("final Auto Pricing workbook")
        if len(base_zips) != 1:
            missing.append("final base-output ZIP")
        if not audit.is_file():
            missing.append("Processing_Audit.xlsx")
        if not report.is_file():
            missing.append("processing_report.json")
        if missing:
            raise RuntimeError(
                "Auto Pricing completed without all expected v4.1.3 outputs: "
                + ", ".join(missing)
                + "\n"
                + logs[-12000:]
            )

        output_zip = work / "Auto_Pricing_v4_1_3_Output.zip"
        with zipfile.ZipFile(output_zip, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.write(final_books[0], arcname=final_books[0].name)
            zf.write(base_zips[0], arcname=base_zips[0].name)
            zf.write(audit, arcname=audit.name)
            zf.write(report, arcname=report.name)

        return output_zip.read_bytes(), output_zip.name, logs
