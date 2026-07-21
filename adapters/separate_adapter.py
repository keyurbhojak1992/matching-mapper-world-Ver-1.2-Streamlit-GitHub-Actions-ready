from __future__ import annotations

import io
import zipfile
from datetime import datetime
from pathlib import Path, PurePosixPath

from openpyxl import load_workbook

from .common import safe_filename, temp_workspace, write_bytes


def _safe_members(zf: zipfile.ZipFile):
    for info in zf.infolist():
        path = PurePosixPath(info.filename)
        if path.is_absolute() or ".." in path.parts:
            raise ValueError(f"Unsafe ZIP entry rejected: {info.filename}")
        yield info


def _unique_name(existing: set[str], candidate: str) -> str:
    path = Path(candidate)
    stem, suffix = path.stem, path.suffix
    current = candidate
    counter = 1
    while current.lower() in existing:
        current = f"{stem}_{counter}{suffix}"
        counter += 1
    existing.add(current.lower())
    return current


def _split_workbook(workbook_path: Path) -> bytes:
    probe = load_workbook(workbook_path, read_only=True, data_only=False, keep_links=True)
    sheet_names = list(probe.sheetnames)
    probe.close()

    nested_buffer = io.BytesIO()
    used_names: set[str] = set()
    with zipfile.ZipFile(nested_buffer, "w", zipfile.ZIP_DEFLATED) as nested_zip:
        for sheet_name in sheet_names:
            # Reloading the complete workbook and deleting other sheets preserves the
            # selected worksheet's styles, hyperlinks, merged cells and drawings much
            # more faithfully than reconstructing cells in a new workbook.
            wb = load_workbook(workbook_path, data_only=False, keep_links=True)
            for other in list(wb.sheetnames):
                if other != sheet_name:
                    wb.remove(wb[other])
            wb.active = 0

            sheet_filename = _unique_name(
                used_names,
                f"{safe_filename(sheet_name, fallback='Sheet')}.xlsx",
            )
            sheet_buffer = io.BytesIO()
            wb.save(sheet_buffer)
            wb.close()
            nested_zip.writestr(sheet_filename, sheet_buffer.getvalue())

    return nested_buffer.getvalue()


def run_separate_excel_files(*, input_zip_bytes: bytes) -> tuple[bytes, str, str]:
    with temp_workspace("matching_mapper_split_") as work:
        input_path = write_bytes(work / "input.zip", input_zip_bytes)
        extract_dir = work / "extract"
        extract_dir.mkdir(parents=True, exist_ok=True)

        with zipfile.ZipFile(input_path, "r") as source_zip:
            for member in _safe_members(source_zip):
                source_zip.extract(member, extract_dir)

        workbook_paths = sorted(
            p for p in extract_dir.rglob("*")
            if p.is_file() and not p.name.startswith("~$") and p.suffix.lower() in {".xlsx", ".xlsm"}
        )
        if not workbook_paths:
            raise ValueError("The uploaded ZIP does not contain any .xlsx or .xlsm workbooks.")

        final_buffer = io.BytesIO()
        used_zip_names: set[str] = set()
        log_lines = [f"Found {len(workbook_paths)} workbook(s)."]
        with zipfile.ZipFile(final_buffer, "w", zipfile.ZIP_DEFLATED) as final_zip:
            for workbook_path in workbook_paths:
                nested_bytes = _split_workbook(workbook_path)
                nested_name = _unique_name(
                    used_zip_names,
                    f"{safe_filename(workbook_path.stem, fallback='Workbook')}.zip",
                )
                final_zip.writestr(nested_name, nested_bytes)
                log_lines.append(f"Separated {workbook_path.name} into {nested_name}")

        output_name = f"Separated_Excel_Sheets_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.zip"
        return final_buffer.getvalue(), output_name, "\n".join(log_lines)
