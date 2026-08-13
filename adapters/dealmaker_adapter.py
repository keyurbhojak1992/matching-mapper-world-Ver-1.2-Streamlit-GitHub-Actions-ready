from __future__ import annotations

import os
import subprocess
import sys
import tempfile
from pathlib import Path

from adapters.common import MASTER_CODE_DIR


PACKAGE_DIR = MASTER_CODE_DIR / "dealmaker_v8_2_2"
RUNNER = PACKAGE_DIR / "Dealmaker_Production_Runner_v6.py"
DEFAULT_CONTROL = PACKAGE_DIR / "Dealmaker_Full_Control_v6.xlsx"
OUTPUT_ZIP_NAME = "Dealmaker_Production_Review_v8.zip"


def bundled_dealmaker_control_bytes() -> bytes:
    """Return the verified control workbook shipped with the preserved package."""
    if not DEFAULT_CONTROL.is_file():
        raise FileNotFoundError("Bundled Dealmaker_Full_Control_v6.xlsx is missing.")
    return DEFAULT_CONTROL.read_bytes()


def run_dealmaker(
    *,
    sales_bytes: bytes,
    actions_bytes: bytes,
    stock_bytes: bytes,
    control_bytes: bytes | None = None,
    chunk_size: int = 10,
    sales_sheet: str = "",
    actions_sheet: str = "",
    stock_sheet: str = "",
    sales_header_row: int | None = None,
    actions_header_row: int | None = None,
    stock_header_row: int | None = None,
) -> tuple[bytes, str, str]:
    """Run the supplied Dealmaker v8.2.2 production runner without modifying package logic."""
    if not RUNNER.is_file() or not DEFAULT_CONTROL.is_file():
        raise FileNotFoundError("The preserved Dealmaker v8.2.2 production package is incomplete.")

    with tempfile.TemporaryDirectory(prefix="matching_mapper_dealmaker_") as tmp:
        work = Path(tmp)
        sales_path = work / "Sales.xlsx"
        actions_path = work / "Actions.xlsx"
        stock_path = work / "Stock.xlsx"
        sales_path.write_bytes(sales_bytes)
        actions_path.write_bytes(actions_bytes)
        stock_path.write_bytes(stock_bytes)

        if control_bytes is None:
            control_path = DEFAULT_CONTROL
            control_source = "Bundled Dealmaker_Full_Control_v6.xlsx"
        else:
            control_path = work / "Dealmaker_Full_Control_v6.xlsx"
            control_path.write_bytes(control_bytes)
            control_source = "Uploaded Dealmaker_Full_Control_v6.xlsx override"

        output_dir = work / "dealmaker_v8_2_output"
        command = [
            sys.executable,
            str(RUNNER),
            "--sales", str(sales_path),
            "--actions", str(actions_path),
            "--stock", str(stock_path),
            "--control", str(control_path),
            "--output-dir", str(output_dir),
            "--chunk-size", str(int(chunk_size)),
        ]

        for flag, value in (
            ("--sales-sheet", sales_sheet),
            ("--actions-sheet", actions_sheet),
            ("--stock-sheet", stock_sheet),
        ):
            if str(value or "").strip():
                command.extend([flag, str(value).strip()])

        for flag, value in (
            ("--sales-header-row", sales_header_row),
            ("--actions-header-row", actions_header_row),
            ("--stock-header-row", stock_header_row),
        ):
            if value is not None:
                command.extend([flag, str(int(value))])

        env = {**os.environ, "PYTHONHASHSEED": "0", "PYTHONUNBUFFERED": "1"}
        process = subprocess.run(
            command,
            cwd=str(PACKAGE_DIR),
            env=env,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            check=False,
        )
        logs = f"Control source: {control_source}\n\n{process.stdout or ''}"
        if process.returncode != 0:
            diagnostic = output_dir / "RUN_FAILURE_DIAGNOSTIC_v8_2_2.txt"
            if diagnostic.is_file():
                logs += "\n\n" + diagnostic.read_text(encoding="utf-8", errors="replace")
            raise RuntimeError("Dealmaker production run failed.\n" + logs[-9000:])

        output_zip = output_dir / OUTPUT_ZIP_NAME
        if not output_zip.is_file():
            raise RuntimeError("Dealmaker completed without creating the production review ZIP.\n" + logs[-9000:])
        return output_zip.read_bytes(), output_zip.name, logs
