from __future__ import annotations

import io
import re
import shutil
import sys
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook

from .common import MASTER_CODE_DIR, captured_output, load_fresh_module, temp_workspace, write_bytes

MASTER_PATH = MASTER_CODE_DIR / "USA_Suggested_Stock_Mapper.py"


def _safe_upload_name(name: str, index: int) -> str:
    base = Path(str(name or "")).name
    stem = re.sub(r"[^A-Za-z0-9._ -]+", "_", base).strip(" .")
    if not stem:
        stem = f"sales_{index}.xlsx"
    if not stem.lower().endswith(".xlsx"):
        stem += ".xlsx"
    return stem


def _worksheet_has_values(worksheet) -> bool:
    for row in worksheet.iter_rows(
        min_row=1,
        max_row=min(worksheet.max_row, 25),
        min_col=1,
        max_col=min(worksheet.max_column, 50),
    ):
        if any(cell.value not in (None, "") for cell in row):
            return True
    return False


def _prepare_stock_workbook(source_path: Path, normalized_path: Path) -> tuple[Path, str]:
    """Ensure the temporary stock workbook has the Sheet1 tab expected by the USA master."""
    try:
        workbook = load_workbook(source_path, read_only=False, data_only=False, keep_links=True)
    except Exception as exc:
        raise ValueError(f"Unable to open the Stock Actual workbook: {exc}") from exc

    try:
        if not workbook.worksheets:
            raise ValueError("The Stock Actual workbook does not contain any worksheets.")
        if "Sheet1" in workbook.sheetnames:
            return source_path, "Sheet1"

        normalized_lookup = {
            re.sub(r"[^a-z0-9]", "", ws.title.casefold()): ws
            for ws in workbook.worksheets
        }
        selected = None
        for preferred in ("stock actual", "stockactual", "stock", "inventory", "data", "sheet 1"):
            selected = normalized_lookup.get(re.sub(r"[^a-z0-9]", "", preferred.casefold()))
            if selected is not None:
                break
        if selected is None:
            active = workbook.active
            if active in workbook.worksheets and _worksheet_has_values(active):
                selected = active
        if selected is None:
            selected = next((ws for ws in workbook.worksheets if _worksheet_has_values(ws)), workbook.worksheets[0])

        original_title = selected.title
        selected.title = "Sheet1"
        workbook.save(normalized_path)
        return normalized_path, original_title
    finally:
        workbook.close()


def run_usa_suggested_stock_mapper(
    *,
    stock_bytes: bytes,
    sales_history_files: Iterable[tuple[str, bytes]],
    top_n_supertags: int = 251,
) -> tuple[bytes, str, str]:
    """Run the supplied USA-team suggested stock mapper with Streamlit-safe uploaded files."""
    sales_history_files = list(sales_history_files)
    if not stock_bytes:
        raise ValueError("Stock Actual workbook is required.")
    if not sales_history_files:
        raise ValueError("At least one USA sales history workbook is required.")
    if int(top_n_supertags) < 1:
        raise ValueError("Top Super Tags must be at least 1.")

    with temp_workspace("matching_mapper_usa_suggested_") as work:
        raw_stock = write_bytes(work / "stock_uploaded.xlsx", stock_bytes)
        stock_path, original_stock_sheet = _prepare_stock_workbook(
            raw_stock,
            work / "stock_actual_usa.xlsx",
        )

        sales_dir = work / "sales_data"
        sales_dir.mkdir(parents=True, exist_ok=True)
        seen_names: set[str] = set()
        for index, (name, data) in enumerate(sales_history_files, start=1):
            safe_name = _safe_upload_name(name, index)
            stem, suffix = Path(safe_name).stem, Path(safe_name).suffix
            candidate = safe_name
            counter = 2
            while candidate.casefold() in seen_names:
                candidate = f"{stem}_{counter}{suffix}"
                counter += 1
            seen_names.add(candidate.casefold())
            write_bytes(sales_dir / candidate, data)

        output_name = f"USA_Suggested_Stock_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.zip"
        output_path = work / output_name

        module = None
        logs_text = ""
        try:
            module = load_fresh_module(MASTER_PATH, "usa_suggested_stock_master")
            # The supplied Colab source imports tqdm.notebook, which requires
            # Jupyter widgets. Streamlit/server execution should use the plain
            # text tqdm implementation instead.
            try:
                from tqdm import tqdm as console_tqdm
                module.tqdm = console_tqdm
            except Exception:
                pass
            with captured_output() as logs:
                if original_stock_sheet != "Sheet1":
                    print(
                        f"Stock worksheet '{original_stock_sheet}' was mapped to temporary 'Sheet1' "
                        "for the USA suggested mapper."
                    )
                module.generate_salesperson_reports(
                    file_path_stock=str(stock_path),
                    sales_files_pattern=str(sales_dir / "*.xlsx"),
                    output_zip_path=str(output_path),
                    top_n=int(top_n_supertags),
                )
                logs_text = logs.getvalue()
        finally:
            if module is not None:
                sys.modules.pop(module.__name__, None)

        if not output_path.exists():
            detail = f"\n\nProcessing log:\n{logs_text}" if logs_text else ""
            raise RuntimeError("USA suggested mapper did not create an output ZIP." + detail)
        if not zipfile.is_zipfile(output_path):
            raise RuntimeError("USA suggested mapper created an invalid ZIP output.")

        with zipfile.ZipFile(output_path, "r") as zf:
            packaged = [name for name in zf.namelist() if not name.endswith("/")]
        if not packaged:
            detail = f"\n\nProcessing log:\n{logs_text}" if logs_text else ""
            raise RuntimeError(
                "No USA party reports were generated. Check that the stock and sales-history "
                "columns/matching values overlap." + detail
            )

        return output_path.read_bytes(), output_name, logs_text
