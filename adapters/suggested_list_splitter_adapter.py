from __future__ import annotations

import io
import shutil
import tempfile
import zipfile
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook


UNDER_10_FOLDER = "Under 10 Stones"
TEN_PLUS_FOLDER = "10+ Stones"
NEEDS_REVIEW_FOLDER = "Needs Review"

SERIAL_HEADERS = {
    "serial no",
    "serial no.",
    "serial number",
    "serial number.",
    "sr no",
    "sr no.",
    "sr. no",
    "sr. no.",
    "sr number",
    "sr. number",
}


def _normalize_header(value: object) -> str:
    if value is None:
        return ""
    return " ".join(str(value).replace("\n", " ").replace("\r", " ").split()).strip().lower()


def _safe_filename(name: str) -> str:
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        name = name.replace(char, "_")
    return name.strip() or "Unknown Salesperson"


def _safe_extract(zf: zipfile.ZipFile, destination: Path) -> None:
    destination_resolved = destination.resolve()
    for member in zf.infolist():
        target = (destination / member.filename).resolve()
        if destination_resolved != target and destination_resolved not in target.parents:
            raise ValueError(f"Unsafe ZIP entry: {member.filename}")
    zf.extractall(destination)


def _remove_noise(base_folder: Path) -> None:
    for path in sorted(base_folder.rglob("*"), key=lambda p: len(p.parts), reverse=True):
        if path.is_file() and path.name in {".DS_Store", "Thumbs.db"}:
            path.unlink(missing_ok=True)
        elif path.is_dir() and path.name == "__MACOSX":
            shutil.rmtree(path, ignore_errors=True)


def _effective_root(base_folder: Path) -> Path:
    current = base_folder
    while True:
        entries = [p for p in current.iterdir() if p.name not in {"__MACOSX", ".DS_Store"}]
        folders = [p for p in entries if p.is_dir()]
        excel_files_here = [p for p in entries if p.is_file() and p.suffix.lower() == ".xlsx"]
        if len(folders) == 1 and not excel_files_here:
            only_folder = folders[0]
            # Treat a sole folder that directly contains XLSX files as the
            # salesperson folder, not as a disposable ZIP wrapper. This fixes
            # the single-salesperson edge case in the original Colab script.
            child_has_excel = any(
                p.is_file() and p.suffix.lower() == ".xlsx"
                for p in only_folder.iterdir()
            )
            if child_has_excel:
                return current
            current = only_folder
            continue
        return current


def _count_stones(excel_file: Path) -> int | None:
    workbook = None
    try:
        workbook = load_workbook(excel_file, read_only=True, data_only=True)
        valid_counts: list[int] = []
        for worksheet in workbook.worksheets:
            serial_column = None
            header_row = None
            max_search_row = min(30, worksheet.max_row)
            for row in worksheet.iter_rows(min_row=1, max_row=max_search_row):
                for cell in row:
                    if _normalize_header(cell.value) in SERIAL_HEADERS:
                        serial_column = cell.column
                        header_row = cell.row
                        break
                if serial_column is not None:
                    break

            if serial_column is None or header_row is None:
                continue

            stone_count = 0
            for row in worksheet.iter_rows(
                min_row=header_row + 1,
                min_col=serial_column,
                max_col=serial_column,
                values_only=True,
            ):
                value = row[0]
                if value is not None and str(value).strip() != "":
                    stone_count += 1
            valid_counts.append(stone_count)

        return max(valid_counts) if valid_counts else None
    except Exception:
        return None
    finally:
        if workbook is not None:
            workbook.close()


def _salesperson_name(excel_file: Path, effective_root: Path) -> str:
    relative = excel_file.relative_to(effective_root)
    return relative.parts[0] if len(relative.parts) >= 2 else "Unknown Salesperson"


def _unique_destination(folder: Path, filename: str) -> Path:
    destination = folder / filename
    if not destination.exists():
        return destination
    stem = Path(filename).stem
    suffix = Path(filename).suffix
    counter = 2
    while True:
        candidate = folder / f"{stem}_{counter}{suffix}"
        if not candidate.exists():
            return candidate
        counter += 1


def _write_nested_zip(source_folder: Path, output_zip: Path) -> None:
    with zipfile.ZipFile(output_zip, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for category in (UNDER_10_FOLDER, TEN_PLUS_FOLDER, NEEDS_REVIEW_FOLDER):
            folder = source_folder / category
            files = sorted(p for p in folder.iterdir() if p.is_file()) if folder.is_dir() else []
            if category != NEEDS_REVIEW_FOLDER or files:
                if not files:
                    zf.writestr(category + "/", "")
                for file_path in files:
                    zf.write(file_path, arcname=f"{category}/{file_path.name}")


def run_suggested_list_splitter(*, input_zip_bytes: bytes) -> tuple[bytes, str, str]:
    """Split suggested XLSX files by stone count and package one ZIP per salesperson.

    The split follows the supplied Colab logic exactly at the boundary:
    fewer than 10 stones -> Under 10 Stones; 10 or more -> 10+ Stones.
    Files whose stone count cannot be determined are preserved in Needs Review.
    """
    logs: list[str] = []
    with tempfile.TemporaryDirectory(prefix="matching_mapper_suggested_list_") as temp_dir:
        work = Path(temp_dir)
        extract_dir = work / "base_files"
        sorted_dir = work / "sorted_output"
        salesperson_zips = work / "salesperson_zips"
        extract_dir.mkdir(parents=True)
        sorted_dir.mkdir(parents=True)
        salesperson_zips.mkdir(parents=True)

        try:
            with zipfile.ZipFile(io.BytesIO(input_zip_bytes), "r") as zf:
                _safe_extract(zf, extract_dir)
        except zipfile.BadZipFile as exc:
            raise ValueError("Uploaded file is not a valid ZIP.") from exc

        _remove_noise(extract_dir)
        effective_root = _effective_root(extract_dir)
        excel_files = sorted(
            p for p in effective_root.rglob("*.xlsx") if not p.name.startswith("~$")
        )
        if not excel_files:
            raise ValueError("No XLSX files found inside the uploaded ZIP.")

        summary: dict[str, dict[str, int]] = defaultdict(
            lambda: {"under": 0, "ten_plus": 0, "review": 0}
        )
        logs.append(f"Found {len(excel_files)} Excel file(s).")

        for index, excel_file in enumerate(excel_files, start=1):
            salesperson = _salesperson_name(excel_file, effective_root)
            salesperson_folder = sorted_dir / salesperson
            under_folder = salesperson_folder / UNDER_10_FOLDER
            ten_plus_folder = salesperson_folder / TEN_PLUS_FOLDER
            review_folder = salesperson_folder / NEEDS_REVIEW_FOLDER
            under_folder.mkdir(parents=True, exist_ok=True)
            ten_plus_folder.mkdir(parents=True, exist_ok=True)

            stones = _count_stones(excel_file)
            if stones is None:
                review_folder.mkdir(parents=True, exist_ok=True)
                destination_folder = review_folder
                summary[salesperson]["review"] += 1
                category = NEEDS_REVIEW_FOLDER
                stone_label = "unknown"
            elif stones < 10:
                destination_folder = under_folder
                summary[salesperson]["under"] += 1
                category = UNDER_10_FOLDER
                stone_label = str(stones)
            else:
                destination_folder = ten_plus_folder
                summary[salesperson]["ten_plus"] += 1
                category = TEN_PLUS_FOLDER
                stone_label = str(stones)

            destination = _unique_destination(destination_folder, excel_file.name)
            shutil.copy2(excel_file, destination)
            logs.append(
                f"[{index}/{len(excel_files)}] {salesperson} | {excel_file.name} | "
                f"stones={stone_label} -> {category}"
            )

        created_zips: list[Path] = []
        for salesperson in sorted(summary):
            salesperson_folder = sorted_dir / salesperson
            zip_path = salesperson_zips / f"{_safe_filename(salesperson)}.zip"
            _write_nested_zip(salesperson_folder, zip_path)
            created_zips.append(zip_path)

        output_name = f"Suggested_List_Split_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.zip"
        output_path = work / output_name
        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as final_zip:
            for zip_path in created_zips:
                final_zip.write(zip_path, arcname=zip_path.name)

        if not zipfile.is_zipfile(output_path):
            raise RuntimeError("Suggested list split completed without a valid output ZIP.")

        total_under = sum(data["under"] for data in summary.values())
        total_ten_plus = sum(data["ten_plus"] for data in summary.values())
        total_review = sum(data["review"] for data in summary.values())
        logs.extend(
            [
                "",
                "Summary",
                f"Under 10 Stones: {total_under}",
                f"10+ Stones: {total_ten_plus}",
                f"Needs Review: {total_review}",
            ]
        )
        return output_path.read_bytes(), output_name, "\n".join(logs)
