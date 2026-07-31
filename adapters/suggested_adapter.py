from __future__ import annotations

import argparse
import gc
import json
import os
import shutil
import subprocess
import sys
import tempfile
import traceback
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable, Iterable, Sequence

import pyarrow.parquet as pq
from openpyxl import load_workbook

from .common import (
    MASTER_CODE_DIR,
    ROOT_DIR,
    apply_runtime_overrides,
    captured_output,
    load_fresh_module,
    resolve_repository_input,
    temp_workspace,
    write_bytes,
)
from .suggested_result_assembler import (
    ASSEMBLER_VERSION,
    DIAGNOSTICS_FOLDER,
    EXCEL_MAX_DATA_ROWS,
    FINAL_MASTER_SHEET_POLICY,
    MASTER_REPORT_NAME,
    add_diagnostic_file,
    collect_preserved_batch_output,
    merge_master_workbooks,
)

MASTER_PATH = MASTER_CODE_DIR / "Suggested_Process.py"
StatusCallback = Callable[[dict[str, object]], None]


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def _normalized_sheet_name(value: object) -> str:
    return "".join(character for character in str(value or "").casefold() if character.isalnum())


def _worksheet_has_values(worksheet) -> bool:
    """Return True when a worksheet contains at least one non-empty cell."""
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
    """
    Ensure the temporary workbook contains the ``Sheet1`` tab required by the
    preserved master process.

    The uploaded workbook itself is never changed. If ``Sheet1`` is missing,
    the most likely stock-data worksheet is renamed only in a temporary copy.
    """
    try:
        workbook = load_workbook(source_path, read_only=False, data_only=False, keep_links=True)
    except Exception as exc:
        raise ValueError(f"Unable to open the Stock Actual workbook as a valid .xlsx file: {exc}") from exc

    try:
        if not workbook.worksheets:
            raise ValueError("The Stock Actual workbook does not contain any worksheets.")

        if "Sheet1" in workbook.sheetnames:
            return source_path, "Sheet1"

        preferred_names = (
            "Stock Actual",
            "StockActual",
            "Stock",
            "Inventory",
            "Data",
            "Sheet 1",
        )
        normalized_lookup = {
            _normalized_sheet_name(worksheet.title): worksheet
            for worksheet in workbook.worksheets
        }

        selected = None
        for preferred_name in preferred_names:
            candidate = normalized_lookup.get(_normalized_sheet_name(preferred_name))
            if candidate is not None:
                selected = candidate
                break

        if selected is None:
            active = workbook.active
            if active in workbook.worksheets and _worksheet_has_values(active):
                selected = active

        if selected is None:
            selected = next(
                (worksheet for worksheet in workbook.worksheets if _worksheet_has_values(worksheet)),
                workbook.worksheets[0],
            )

        original_title = selected.title
        selected.title = "Sheet1"
        workbook.save(normalized_path)
        return normalized_path, original_title
    except ValueError:
        raise
    except Exception as exc:
        available = ", ".join(workbook.sheetnames) or "none"
        raise ValueError(
            "Unable to prepare the Stock Actual workbook for processing. "
            f"Available worksheets: {available}. Details: {exc}"
        ) from exc
    finally:
        workbook.close()


def _copy_optional(source: Path | None, destination: Path) -> Path:
    """Copy an optional input when present, while returning the expected path."""
    if source is not None and source.is_file():
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, destination)
    return destination


def _execute_master_in_workspace(
    *,
    work: Path,
    stock_path: Path,
    profile_perm_path: Path,
    profile_temp_path: Path,
    filter_path: Path,
    clean_path: Path,
    sales_dir: Path,
    runtime_overrides: dict[str, object] | None,
) -> tuple[Path, str]:
    normalized_stock_path, original_stock_sheet = _prepare_stock_workbook(
        stock_path,
        work / "stock Actual.xlsx",
    )
    output_name = f"Matching_Mapper_Suggested_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.zip"
    output_path = work / output_name

    module = None
    logs_text = ""
    try:
        module = load_fresh_module(MASTER_PATH, "suggested_master")
        apply_runtime_overrides(module, runtime_overrides)

        # Rebuild constants that the untouched source computes at import time.
        if hasattr(module, "ENABLE_FLO_MAPPING"):
            module.FLUORESCENCE_ORDER = (
                ["NON", "FNT", "MED", "STG", "VST"]
                if module.ENABLE_FLO_MAPPING
                else ["NON", "VSL", "SL", "FNT", "MED", "STG", "VST"]
            )
            if hasattr(module, "ORDER_MAPPING"):
                module.ORDER_MAPPING["FLO"] = module.FLUORESCENCE_ORDER
                module.ORDER_MAPPING["Flo"] = module.FLUORESCENCE_ORDER

        with captured_output() as logs:
            try:
                if original_stock_sheet != "Sheet1":
                    print(
                        "Stock workbook worksheet "
                        f"'{original_stock_sheet}' was mapped to temporary 'Sheet1' for the preserved master process."
                    )
                module.generate_salesperson_reports_from_profile(
                    str(normalized_stock_path),
                    str(profile_perm_path),
                    str(profile_temp_path),
                    str(output_path),
                    str(filter_path),
                    str(clean_path),
                    str(sales_dir),
                    getattr(module, "TOP_N_SUPERTAGS", None),
                    int(getattr(module, "TOLERANCE_LEVEL", 1)),
                )
            finally:
                logs_text = logs.getvalue()

        if not output_path.exists():
            raise RuntimeError(
                "The master process completed without creating an output ZIP. "
                "Review the processing log and input profiles/filters.\n\n" + logs_text
            )
        return output_path, logs_text
    except Exception as exc:
        if logs_text:
            raise RuntimeError(f"{exc}\n\nCaptured processing output:\n{logs_text}") from exc
        raise
    finally:
        if module is not None:
            sys.modules.pop(module.__name__, None)
            del module
        gc.collect()


def _materialize_suggested_input(
    *,
    work: Path,
    uploaded_bytes: bytes | None,
    default_key: str,
    destination_name: str,
    required: bool = False,
) -> Path:
    destination = work / destination_name
    if uploaded_bytes is not None:
        return write_bytes(destination, uploaded_bytes)
    source = resolve_repository_input(None, default_key, required=required)
    if source is not None:
        shutil.copy2(source, destination)
    return destination


def run_suggested_process(
    *,
    stock_bytes: bytes,
    profile_perm_bytes: bytes | None,
    profile_temp_bytes: bytes | None,
    filter_bytes: bytes | None,
    clean_bytes: bytes | None,
    sales_history_files: Iterable[tuple[str, bytes]] = (),
    runtime_overrides: dict[str, object] | None = None,
) -> tuple[bytes, str, str]:
    """Execute Suggested Process, using repository defaults for omitted files."""
    with temp_workspace("matching_mapper_suggested_") as work:
        stock_path = write_bytes(work / "stock_uploaded.xlsx", stock_bytes)
        filter_path = _materialize_suggested_input(
            work=work, uploaded_bytes=filter_bytes, default_key="filters",
            destination_name="filters.xlsx", required=True,
        )
        perm_path = _materialize_suggested_input(
            work=work, uploaded_bytes=profile_perm_bytes, default_key="profile_perm",
            destination_name="PARTY_PROFILE_PERM.parquet",
        )
        temp_path = _materialize_suggested_input(
            work=work, uploaded_bytes=profile_temp_bytes, default_key="profile_temp",
            destination_name="PARTY_PROFILE_TEMP.parquet",
        )
        clean_path = _materialize_suggested_input(
            work=work, uploaded_bytes=clean_bytes, default_key="no_need",
            destination_name="no_need.xlsx",
        )

        sales_dir = work / "Sales_And_Actions_Folder"
        sales_dir.mkdir(parents=True, exist_ok=True)
        for index, (name, data) in enumerate(sales_history_files, start=1):
            suffix = Path(name).suffix.lower() or ".xlsx"
            write_bytes(sales_dir / f"sales_history_{index}{suffix}", data)

        output_path, logs = _execute_master_in_workspace(
            work=work,
            stock_path=stock_path,
            profile_perm_path=perm_path,
            profile_temp_path=temp_path,
            filter_path=filter_path,
            clean_path=clean_path,
            sales_dir=sales_dir,
            runtime_overrides=runtime_overrides,
        )
        return output_path.read_bytes(), output_path.name, logs


def run_suggested_process_from_paths(
    *,
    stock_path: Path,
    profile_perm_path: Path | None,
    profile_temp_path: Path | None,
    filter_path: Path | None,
    clean_path: Path | None,
    sales_history_paths: Iterable[Path] = (),
    runtime_overrides: dict[str, object] | None = None,
    output_path: Path | None = None,
) -> tuple[Path, str]:
    """Run one Suggested batch, resolving omitted inputs from ``data``."""
    stock_path = Path(stock_path)
    if not stock_path.is_file():
        raise FileNotFoundError(f"Stock workbook not found: {stock_path}")

    resolved_filter = resolve_repository_input(filter_path, "filters", required=True)
    resolved_perm = resolve_repository_input(profile_perm_path, "profile_perm")
    resolved_temp = resolve_repository_input(profile_temp_path, "profile_temp")
    resolved_clean = resolve_repository_input(clean_path, "no_need")
    assert resolved_filter is not None

    with temp_workspace("matching_mapper_suggested_batch_") as work:
        local_stock = work / "stock_uploaded.xlsx"
        local_filter = work / "filters.xlsx"
        shutil.copy2(stock_path, local_stock)
        shutil.copy2(resolved_filter, local_filter)

        local_perm = _copy_optional(resolved_perm, work / "PARTY_PROFILE_PERM.parquet")
        local_temp = _copy_optional(resolved_temp, work / "PARTY_PROFILE_TEMP.parquet")
        local_clean = _copy_optional(resolved_clean, work / "no_need.xlsx")

        sales_dir = work / "Sales_And_Actions_Folder"
        sales_dir.mkdir(parents=True, exist_ok=True)
        for index, source in enumerate(sales_history_paths, start=1):
            source = Path(source)
            if not source.is_file():
                continue
            suffix = source.suffix.lower() or ".xlsx"
            shutil.copy2(source, sales_dir / f"sales_history_{index}{suffix}")

        generated_path, logs = _execute_master_in_workspace(
            work=work,
            stock_path=local_stock,
            profile_perm_path=local_perm,
            profile_temp_path=local_temp,
            filter_path=local_filter,
            clean_path=local_clean,
            sales_dir=sales_dir,
            runtime_overrides=runtime_overrides,
        )

        destination = (
            Path.cwd() / f".suggested_batch_{datetime.now().strftime('%Y%m%d%H%M%S%f')}.zip"
            if output_path is None else Path(output_path)
        )
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(generated_path, destination)
        return destination, logs

def _clean_values(values: Iterable[object]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        text = str(value).strip() if value is not None else ""
        if not text or text.casefold() in {"nan", "none"}:
            continue
        key = text.casefold()
        if key not in seen:
            seen.add(key)
            result.append(text)
    return result


def _normalized_column(value: object) -> str:
    return "".join(character for character in str(value or "").casefold() if character.isalnum())


def _parquet_salespeople(path: Path | None) -> list[str]:
    if path is None or not Path(path).is_file():
        return []
    parquet = pq.ParquetFile(path)
    candidates = {"salesperson", "salespersonname"}
    matching = [name for name in parquet.schema_arrow.names if _normalized_column(name) in candidates]
    if not matching:
        return []
    table = parquet.read(columns=matching)
    values: list[object] = []
    for name in matching:
        values.extend(table.column(name).to_pylist())
    return _clean_values(values)


def _filter_salespeople(path: Path | None) -> list[str]:
    if path is None or not Path(path).is_file():
        return []
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        preferred = [
            worksheet
            for worksheet in workbook.worksheets
            if _normalized_sheet_name(worksheet.title) in {"changeformat", "tolerance"}
        ]
        worksheets = preferred or list(workbook.worksheets)
        values: list[str] = []
        for worksheet in worksheets:
            header_values = next(worksheet.iter_rows(min_row=1, max_row=1, values_only=True), ())
            indexes = [
                index
                for index, header in enumerate(header_values)
                if _normalized_column(header) == "salesperson"
            ]
            if not indexes:
                continue
            for row in worksheet.iter_rows(min_row=2, values_only=True):
                for index in indexes:
                    if index < len(row) and row[index] not in (None, ""):
                        # The preserved filter loader normalizes this column with pandas .str.title().
                        values.append(str(row[index]).strip().title())
        return _clean_values(values)
    finally:
        workbook.close()


def discover_salespeople_from_paths(
    *,
    profile_perm_path: Path | None,
    profile_temp_path: Path | None,
    filter_path: Path | None,
) -> list[str]:
    """Read only salesperson columns from the profiles and filter workbook."""
    values = (
        _parquet_salespeople(profile_perm_path)
        + _parquet_salespeople(profile_temp_path)
        + _filter_salespeople(filter_path)
    )
    return sorted(_clean_values(values), key=str.casefold)


def _json_safe(value: object) -> object:
    if isinstance(value, dict):
        return {str(key): _json_safe(item) for key, item in value.items()}
    if isinstance(value, (set, tuple, list)):
        return [_json_safe(item) for item in value]
    if isinstance(value, Path):
        return str(value)
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


def _chunked(values: Sequence[str], size: int) -> list[list[str]]:
    return [list(values[index:index + size]) for index in range(0, len(values), size)]


def _write_json(path: Path, payload: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(_json_safe(payload), indent=2, ensure_ascii=False), encoding="utf-8")
    temporary.replace(path)


def _run_batch_subprocess(*, spec_path: Path, console_path: Path) -> subprocess.CompletedProcess[bytes]:
    environment = os.environ.copy()
    existing_pythonpath = environment.get("PYTHONPATH", "")
    environment["PYTHONPATH"] = str(ROOT_DIR) + (os.pathsep + existing_pythonpath if existing_pythonpath else "")
    command = [sys.executable, "-m", "adapters.suggested_adapter", "--run-batch-spec", str(spec_path)]
    with console_path.open("wb") as console:
        return subprocess.run(
            command,
            cwd=str(ROOT_DIR),
            env=environment,
            stdout=console,
            stderr=subprocess.STDOUT,
            check=False,
        )


def run_suggested_batched_from_paths(
    *,
    stock_path: Path,
    profile_perm_path: Path | None,
    profile_temp_path: Path | None,
    filter_path: Path | None,
    clean_path: Path | None,
    sales_history_paths: Iterable[Path],
    runtime_overrides: dict[str, object],
    scope: str,
    salespeople: Sequence[str],
    batch_size: int,
    output_path: Path,
    continue_after_failed_batch: bool = True,
    status_callback: StatusCallback | None = None,
    manifest_path: Path | None = None,
    processing_log_path: Path | None = None,
) -> dict[str, object]:
    """
    Run the preserved engine in isolated salesperson batches, then assemble
    the same final layout that an unbatched master run would produce.

    Every batch starts a fresh Python subprocess. This gives the operating
    system a chance to reclaim pandas, PyArrow and openpyxl memory before the
    next group begins. Only ``TEST_MODE`` and ``TEST_SALESPEOPLE`` are changed
    per batch; all other runtime choices are copied unchanged.

    The preserved source files are never edited. Their batch-level outputs are
    post-processed by the adapter as follows:

    * batch master workbooks become one global ``MASTER_CONSOLIDATED_REPORT.xlsx``;
    * all salesperson workbooks are placed together at the final ZIP root;
    * all ``*_Separated_Parties.zip`` files are placed at the final ZIP root;
    * batch logs and runtime snapshots are retained under ``_batch_diagnostics``.
    """
    normalized_scope = str(scope).strip().lower()
    if normalized_scope not in {"all", "selected"}:
        raise ValueError("Batch scope must be 'all' or 'selected'.")
    if not 1 <= int(batch_size) <= 20:
        raise ValueError("Batch size must be between 1 and 20.")

    profile_perm_path = resolve_repository_input(profile_perm_path, "profile_perm")
    profile_temp_path = resolve_repository_input(profile_temp_path, "profile_temp")
    filter_path = resolve_repository_input(filter_path, "filters", required=True)
    clean_path = resolve_repository_input(clean_path, "no_need")
    assert filter_path is not None

    if normalized_scope == "all":
        clean_salespeople = discover_salespeople_from_paths(
            profile_perm_path=profile_perm_path,
            profile_temp_path=profile_temp_path,
            filter_path=filter_path,
        )
    else:
        clean_salespeople = _clean_values(salespeople)
    if not clean_salespeople:
        raise ValueError(
            "No salespeople were found. Check the Salesperson profile column and the SALES_PERSON filter column."
        )

    history_paths = [Path(path) for path in sales_history_paths if Path(path).is_file()]
    batches = _chunked(clean_salespeople, int(batch_size))
    started_at = _utc_now()
    merged_master_requested = bool(runtime_overrides.get("GENERATE_MERGED_MASTER", True))
    individual_reports_requested = bool(runtime_overrides.get("GENERATE_INDIVIDUAL_REPORTS", False))
    separated_reports_requested = bool(runtime_overrides.get("ENABLE_OUTPUT_FILE_SEPARATED", False))
    manifest: dict[str, object] = {
        "schema_version": 3,
        "started_at": started_at,
        "completed_at": None,
        "scope": normalized_scope,
        "batch_size": int(batch_size),
        "total_salespeople": len(clean_salespeople),
        "salespeople": list(clean_salespeople),
        "total_batches": len(batches),
        "continue_after_failed_batch": bool(continue_after_failed_batch),
        "base_runtime": _json_safe(runtime_overrides),
        "batches": [],
        "summary": {"successful": 0, "failed": 0, "skipped": 0},
        "output_layout": {
            "mode": "combined_master_layout",
            "merged_master_requested": merged_master_requested,
            "individual_reports_requested": individual_reports_requested,
            "separated_party_zips_requested": separated_reports_requested,
            "diagnostics_folder": DIAGNOSTICS_FOLDER,
            "root_files": [],
            "master_merge": None,
        },
        "final_state": "running",
    }
    root_log: list[str] = [
        f"[{started_at}] Matching Mapper cloud batch run started.",
        f"Scope: {normalized_scope}",
        f"Salespeople: {len(clean_salespeople)}",
        f"Batches: {len(batches)} (size {batch_size})",
        "Each batch runs in a fresh subprocess.",
        "Final packaging will restore the normal unbatched output layout.",
    ]

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.unlink(missing_ok=True)
    temporary_output_path = output_path.with_suffix(output_path.suffix + ".tmp")
    temporary_output_path.unlink(missing_ok=True)

    try:
        with tempfile.TemporaryDirectory(prefix="matching_mapper_batches_") as temporary_directory:
            batch_root = Path(temporary_directory)
            master_parts_directory = batch_root / "master_parts"
            master_parts: list[Path] = []
            packaged_root_files: list[str] = []
            used_root_names = {
                MASTER_REPORT_NAME.casefold(),
                "batch_manifest.json",
                "processing_log.txt",
                "final_output_readme.txt",
            }

            # XLSX files and nested ZIPs are already compressed, so the final
            # container stores them without spending time recompressing them.
            with zipfile.ZipFile(
                temporary_output_path,
                "w",
                compression=zipfile.ZIP_STORED,
                allowZip64=True,
            ) as final_zip:
                for batch_index, current_batch in enumerate(batches, start=1):
                    batch_name = f"batch_{batch_index:02d}"
                    batch_dir = batch_root / batch_name
                    batch_dir.mkdir(parents=True, exist_ok=True)
                    batch_output = batch_dir / "suggested_reports.zip"
                    batch_log = batch_dir / "processing_log.txt"
                    batch_console = batch_dir / "worker_console.txt"
                    batch_runtime_path = batch_dir / "batch_runtime.json"
                    spec_path = batch_dir / "batch_spec.json"

                    batch_runtime = dict(runtime_overrides)
                    batch_runtime["TEST_MODE"] = True
                    batch_runtime["TEST_SALESPEOPLE"] = list(current_batch)
                    _write_json(batch_runtime_path, batch_runtime)

                    batch_record: dict[str, object] = {
                        "batch": batch_index,
                        "folder": batch_name,
                        "salespeople": list(current_batch),
                        "status": "running",
                        "started_at": _utc_now(),
                        "completed_at": None,
                        "error": None,
                        "packaged_files": [],
                        "master_parts": 0,
                    }
                    manifest_batches = manifest["batches"]
                    assert isinstance(manifest_batches, list)
                    manifest_batches.append(batch_record)

                    if status_callback:
                        status_callback(
                            {
                                "state": "running",
                                "stage": "processing",
                                "message": f"Processing batch {batch_index} of {len(batches)}",
                                "current_batch": batch_index,
                                "total_batches": len(batches),
                                "current_salespeople": list(current_batch),
                                "total_salespeople": len(clean_salespeople),
                                "successful_batches": manifest["summary"]["successful"],  # type: ignore[index]
                                "failed_batches": manifest["summary"]["failed"],  # type: ignore[index]
                                "started_at": started_at,
                                "updated_at": _utc_now(),
                            }
                        )

                    spec = {
                        "stock_path": str(Path(stock_path).resolve()),
                        "profile_perm_path": str(Path(profile_perm_path).resolve()) if profile_perm_path else None,
                        "profile_temp_path": str(Path(profile_temp_path).resolve()) if profile_temp_path else None,
                        "filter_path": str(Path(filter_path).resolve()),
                        "clean_path": str(Path(clean_path).resolve()) if clean_path else None,
                        "sales_history_paths": [str(path.resolve()) for path in history_paths],
                        "runtime_overrides": _json_safe(batch_runtime),
                        "output_path": str(batch_output.resolve()),
                        "log_path": str(batch_log.resolve()),
                    }
                    _write_json(spec_path, spec)

                    try:
                        completed = _run_batch_subprocess(spec_path=spec_path, console_path=batch_console)
                        if completed.returncode != 0:
                            error_detail = (
                                batch_log.read_text(encoding="utf-8", errors="replace")
                                if batch_log.exists()
                                else ""
                            )
                            if not error_detail and batch_console.exists():
                                error_detail = batch_console.read_text(encoding="utf-8", errors="replace")
                            raise RuntimeError(
                                f"Batch subprocess exited with code {completed.returncode}.\n{error_detail[-12000:]}"
                            )
                        if not batch_output.is_file():
                            raise RuntimeError("Batch subprocess succeeded without creating suggested_reports.zip.")

                        collected = collect_preserved_batch_output(
                            batch_zip_path=batch_output,
                            batch_name=batch_name,
                            master_parts_directory=master_parts_directory,
                            final_archive=final_zip,
                            used_root_names=used_root_names,
                        )
                        master_parts.extend(collected.master_parts)
                        packaged_root_files.extend(collected.packaged_files)
                        batch_record["packaged_files"] = collected.packaged_files
                        batch_record["master_parts"] = len(collected.master_parts)
                        batch_record["empty_output"] = collected.empty

                        add_diagnostic_file(
                            final_zip,
                            source_path=batch_log,
                            batch_name=batch_name,
                            filename="processing_log.txt",
                        )
                        add_diagnostic_file(
                            final_zip,
                            source_path=batch_runtime_path,
                            batch_name=batch_name,
                            filename="batch_runtime.json",
                        )
                        if batch_console.exists() and batch_console.stat().st_size:
                            add_diagnostic_file(
                                final_zip,
                                source_path=batch_console,
                                batch_name=batch_name,
                                filename="worker_console.txt",
                            )

                        batch_record["status"] = "success"
                        manifest["summary"]["successful"] += 1  # type: ignore[index]
                        root_log.append(f"[{_utc_now()}] {batch_name}: success ({', '.join(current_batch)})")
                    except Exception as exc:
                        failure_text = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
                        if batch_log.exists():
                            existing = batch_log.read_text(encoding="utf-8", errors="replace")
                            failure_text = existing + "\n\nParent worker error:\n" + failure_text
                        batch_log.write_text(failure_text, encoding="utf-8")
                        add_diagnostic_file(
                            final_zip,
                            source_path=batch_log,
                            batch_name=batch_name,
                            filename="processing_log.txt",
                        )
                        add_diagnostic_file(
                            final_zip,
                            source_path=batch_runtime_path,
                            batch_name=batch_name,
                            filename="batch_runtime.json",
                        )
                        if batch_console.exists() and batch_console.stat().st_size:
                            add_diagnostic_file(
                                final_zip,
                                source_path=batch_console,
                                batch_name=batch_name,
                                filename="worker_console.txt",
                            )
                        batch_record["status"] = "failed"
                        batch_record["error"] = str(exc)
                        manifest["summary"]["failed"] += 1  # type: ignore[index]
                        root_log.append(f"[{_utc_now()}] {batch_name}: failed - {exc}")

                        if not continue_after_failed_batch:
                            remaining = len(batches) - batch_index
                            manifest["summary"]["skipped"] = remaining  # type: ignore[index]
                            for skipped_index in range(batch_index + 1, len(batches) + 1):
                                skipped_people = batches[skipped_index - 1]
                                manifest_batches.append(
                                    {
                                        "batch": skipped_index,
                                        "folder": f"batch_{skipped_index:02d}",
                                        "salespeople": skipped_people,
                                        "status": "skipped",
                                        "started_at": None,
                                        "completed_at": None,
                                        "error": "Skipped after an earlier batch failure.",
                                        "packaged_files": [],
                                        "master_parts": 0,
                                    }
                                )
                            root_log.append(f"Stopped after failure; skipped {remaining} batch(es).")
                            break
                    finally:
                        batch_record["completed_at"] = _utc_now()
                        for disposable in (batch_output, spec_path):
                            disposable.unlink(missing_ok=True)
                        gc.collect()

                if status_callback:
                    status_callback(
                        {
                            "state": "running",
                            "stage": "packaging",
                            "message": "All available batches are processed; assembling the final combined output.",
                            "current_batch": len(batches),
                            "total_batches": len(batches),
                            "total_salespeople": len(clean_salespeople),
                            "successful_batches": manifest["summary"]["successful"],  # type: ignore[index]
                            "failed_batches": manifest["summary"]["failed"],  # type: ignore[index]
                            "skipped_batches": manifest["summary"]["skipped"],  # type: ignore[index]
                            "started_at": started_at,
                            "updated_at": _utc_now(),
                        }
                    )

                output_layout = manifest["output_layout"]
                assert isinstance(output_layout, dict)
                if merged_master_requested:
                    merged_master_path = batch_root / MASTER_REPORT_NAME
                    merge_summary = merge_master_workbooks(
                        master_parts=master_parts,
                        output_path=merged_master_path,
                        # The preserved engine may split intermediate batch workbooks,
                        # but the final user-facing master is always one worksheet.
                        max_rows_per_sheet=EXCEL_MAX_DATA_ROWS,
                        enable_styling=bool(runtime_overrides.get("ENABLE_STYLING", True)),
                        enable_new_fill_logic=bool(runtime_overrides.get("ENABLE_NEW_FILL_LOGIC", True)),
                    )
                    if merge_summary is not None:
                        final_zip.write(
                            merge_summary.output_path,
                            MASTER_REPORT_NAME,
                            compress_type=zipfile.ZIP_STORED,
                        )
                        packaged_root_files.insert(0, MASTER_REPORT_NAME)
                        output_layout["master_merge"] = {
                            "created": True,
                            "source_workbooks": merge_summary.source_workbooks,
                            "rows": merge_summary.total_rows,
                            "columns": merge_summary.total_columns,
                            "worksheets": merge_summary.worksheets,
                            "hyperlinks": merge_summary.hyperlinks,
                            "assembler_version": ASSEMBLER_VERSION,
                            "sheet_policy": FINAL_MASTER_SHEET_POLICY,
                            "sheet_name": "Master_Data",
                        }
                        root_log.append(
                            "Merged master created from "
                            f"{merge_summary.source_workbooks} batch workbook(s): "
                            f"{merge_summary.total_rows:,} rows, {merge_summary.total_columns} columns, "
                            f"{merge_summary.worksheets} worksheet(s), "
                            f"{merge_summary.hyperlinks:,} hyperlinks."
                        )
                    else:
                        output_layout["master_merge"] = {
                            "created": False,
                            "source_workbooks": len(master_parts),
                            "reason": "No non-empty batch master workbook was available.",
                        }
                        root_log.append("No merged master was created because every successful batch was empty.")
                else:
                    output_layout["master_merge"] = {
                        "created": False,
                        "reason": "GENERATE_MERGED_MASTER was disabled.",
                    }

                output_layout["root_files"] = list(packaged_root_files)
                output_layout["salesperson_workbooks"] = len(
                    [
                        name
                        for name in packaged_root_files
                        if name.lower().endswith(".xlsx") and name.casefold() != MASTER_REPORT_NAME.casefold()
                    ]
                )
                output_layout["separated_party_zips"] = len(
                    [name for name in packaged_root_files if name.lower().endswith("_separated_parties.zip")]
                )

                successful = int(manifest["summary"]["successful"])  # type: ignore[index]
                failed = int(manifest["summary"]["failed"])  # type: ignore[index]
                skipped = int(manifest["summary"]["skipped"])  # type: ignore[index]
                if failed == 0 and skipped == 0:
                    final_state = "completed"
                elif successful > 0:
                    final_state = "partial_success"
                else:
                    final_state = "failed"
                manifest["final_state"] = final_state
                manifest["completed_at"] = _utc_now()
                root_log.append(f"[{manifest['completed_at']}] Finished with state: {final_state}.")

                final_readme = (
                    "Matching Mapper World - Final Combined Suggested Output\n"
                    "========================================================\n\n"
                    "This archive restores the normal master-code output layout after low-memory batch processing.\n\n"
                    f"Adapter assembler version: {ASSEMBLER_VERSION}\n"
                    "Final master worksheet policy: strict single sheet (Master_Data).\n\n"
                    f"- {MASTER_REPORT_NAME}: one global merged workbook with one worksheet named Master_Data when Merged Master was enabled and the data fits Excel's worksheet limit.\n"
                    "- <Salesperson>.xlsx: all salesperson workbooks together at this ZIP root when Individual Reports was enabled.\n"
                    "- <Salesperson>_Separated_Parties.zip: all separated-party packages together at this ZIP root when enabled.\n"
                    f"- {DIAGNOSTICS_FOLDER}/: per-batch logs and runtime snapshots for troubleshooting.\n"
                    "- batch_manifest.json and processing_log.txt: complete job summary.\n\n"
                    "The four files in original_master_code remain unchanged; final merging is performed by the adapter wrapper.\n"
                )
                final_zip.writestr(
                    "FINAL_OUTPUT_README.txt",
                    final_readme,
                    compress_type=zipfile.ZIP_DEFLATED,
                )
                root_log_text = "\n".join(root_log) + "\n"
                manifest_text = json.dumps(_json_safe(manifest), indent=2, ensure_ascii=False)
                final_zip.writestr("batch_manifest.json", manifest_text, compress_type=zipfile.ZIP_DEFLATED)
                final_zip.writestr("processing_log.txt", root_log_text, compress_type=zipfile.ZIP_DEFLATED)

        temporary_output_path.replace(output_path)
    except Exception:
        temporary_output_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)
        raise

    if manifest_path is not None:
        Path(manifest_path).write_text(
            json.dumps(_json_safe(manifest), indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
    if processing_log_path is not None:
        Path(processing_log_path).write_text("\n".join(root_log) + "\n", encoding="utf-8")

    if status_callback:
        status_callback(
            {
                "state": manifest["final_state"],
                "stage": "completed",
                "message": "Batch processing and final combined packaging finished.",
                "current_batch": len(batches),
                "total_batches": len(batches),
                "total_salespeople": len(clean_salespeople),
                "successful_batches": manifest["summary"]["successful"],  # type: ignore[index]
                "failed_batches": manifest["summary"]["failed"],  # type: ignore[index]
                "skipped_batches": manifest["summary"]["skipped"],  # type: ignore[index]
                "result_filename": output_path.name,
                "started_at": started_at,
                "completed_at": manifest["completed_at"],
                "updated_at": _utc_now(),
            }
        )
    return manifest

def _run_batch_spec(spec_path: Path) -> int:
    """Internal child-process entry point for one isolated batch."""
    spec = json.loads(Path(spec_path).read_text(encoding="utf-8"))
    output_path = Path(spec["output_path"])
    log_path = Path(spec["log_path"])
    log_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        _, logs = run_suggested_process_from_paths(
            stock_path=Path(spec["stock_path"]),
            profile_perm_path=Path(spec["profile_perm_path"]) if spec.get("profile_perm_path") else None,
            profile_temp_path=Path(spec["profile_temp_path"]) if spec.get("profile_temp_path") else None,
            filter_path=Path(spec["filter_path"]),
            clean_path=Path(spec["clean_path"]) if spec.get("clean_path") else None,
            sales_history_paths=[Path(path) for path in spec.get("sales_history_paths", [])],
            runtime_overrides=dict(spec.get("runtime_overrides") or {}),
            output_path=output_path,
        )
        log_path.write_text(logs or "Process completed.\n", encoding="utf-8")
        return 0
    except Exception as exc:
        log_path.write_text(
            "".join(traceback.format_exception(type(exc), exc, exc.__traceback__)),
            encoding="utf-8",
        )
        return 1


def _main() -> int:
    parser = argparse.ArgumentParser(description="Matching Mapper Suggested adapter worker")
    parser.add_argument("--run-batch-spec", type=Path, help="Internal JSON batch specification")
    arguments = parser.parse_args()
    if arguments.run_batch_spec:
        return _run_batch_spec(arguments.run_batch_spec)
    parser.error("No operation was supplied.")
    return 2


if __name__ == "__main__":
    raise SystemExit(_main())
