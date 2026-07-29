from __future__ import annotations

import json
import shutil
import uuid
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import BinaryIO, Iterable, Mapping, Sequence

SCHEMA_VERSION = 2
INPUT_ASSET_NAME = "input_bundle.zip"
RESULT_ASSET_NAME = "suggested_reports_final_combined.zip"
LEGACY_RESULT_ASSET_NAME = "suggested_reports_all_batches.zip"
STATUS_ASSET_NAME = "job_status.json"
MANIFEST_ASSET_NAME = "batch_manifest.json"
LOG_ASSET_NAME = "processing_log.txt"
CONSOLE_ASSET_NAME = "worker_console.txt"
FAILURE_ASSET_NAME = "failure_log.txt"

# GitHub currently requires each release asset to be smaller than 2 GiB.
# Leave headroom for protocol and platform differences.
MAX_RELEASE_ASSET_BYTES = 1_950_000_000


def new_job_tag() -> str:
    """Return a unique, human-readable identifier for one cloud job."""
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")
    return f"mmw-job-{timestamp}-{uuid.uuid4().hex[:8]}"


def json_safe(value: object) -> object:
    """Convert runtime values such as sets into JSON-compatible objects."""
    if isinstance(value, dict):
        return {str(key): json_safe(item) for key, item in value.items()}
    if isinstance(value, (set, tuple, list)):
        return [json_safe(item) for item in value]
    if isinstance(value, Path):
        return str(value)
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


def _clean_original_name(name: str, fallback: str) -> str:
    text = Path(str(name or fallback)).name.strip()
    return text or fallback


def _rewind(source: BinaryIO) -> None:
    try:
        source.seek(0)
    except (AttributeError, OSError):
        pass


def _write_stream(zf: zipfile.ZipFile, arcname: str, source: BinaryIO) -> None:
    """Copy a file-like object into a ZIP without first materializing another bytes copy."""
    _rewind(source)
    info = zipfile.ZipInfo(arcname)
    # XLSX and Parquet are already compressed; storing avoids wasting Streamlit CPU.
    info.compress_type = zipfile.ZIP_STORED
    info.external_attr = 0o600 << 16
    with zf.open(info, mode="w", force_zip64=True) as destination:
        shutil.copyfileobj(source, destination, length=1024 * 1024)
    _rewind(source)


def _dedupe_selected(values: Sequence[str]) -> list[str]:
    output: list[str] = []
    seen: set[str] = set()
    for value in values:
        text = str(value).strip()
        key = text.casefold()
        if text and key not in seen:
            seen.add(key)
            output.append(text)
    return output


def build_job_bundle(
    *,
    output_path: str | Path,
    job_id: str,
    runtime_overrides: Mapping[str, object],
    scope: str,
    salespeople: Sequence[str],
    batch_size: int,
    continue_after_failed_batch: bool,
    stock_file: tuple[str, BinaryIO],
    filter_file: tuple[str, BinaryIO],
    profile_perm_file: tuple[str, BinaryIO] | None = None,
    profile_temp_file: tuple[str, BinaryIO] | None = None,
    clean_file: tuple[str, BinaryIO] | None = None,
    sales_history_files: Iterable[tuple[str, BinaryIO]] = (),
    requested_by: str = "",
) -> dict[str, object]:
    """Build the disk-backed ZIP uploaded to a private GitHub draft release."""
    normalized_scope = str(scope).strip().lower()
    if normalized_scope not in {"all", "selected"}:
        raise ValueError("Cloud scope must be either 'all' or 'selected'.")

    cleaned_salespeople = _dedupe_selected(salespeople)
    if normalized_scope == "selected" and not cleaned_salespeople:
        raise ValueError("At least one salesperson is required for selected-scope processing.")
    if not 1 <= int(batch_size) <= 20:
        raise ValueError("Batch size must be between 1 and 20.")

    history_items = list(sales_history_files)
    inputs: dict[str, object] = {
        "stock": {
            "path": "inputs/stock_actual.xlsx",
            "original_name": _clean_original_name(stock_file[0], "stock_actual.xlsx"),
        },
        "filters": {
            "path": "inputs/filters.xlsx",
            "original_name": _clean_original_name(filter_file[0], "filters.xlsx"),
        },
        "profile_perm": None,
        "profile_temp": None,
        "clean": None,
        "sales_history": [],
    }

    if profile_perm_file:
        inputs["profile_perm"] = {
            "path": "inputs/PARTY_PROFILE_PERM.parquet",
            "original_name": _clean_original_name(profile_perm_file[0], "PARTY_PROFILE_PERM.parquet"),
        }
    if profile_temp_file:
        inputs["profile_temp"] = {
            "path": "inputs/PARTY_PROFILE_TEMP.parquet",
            "original_name": _clean_original_name(profile_temp_file[0], "PARTY_PROFILE_TEMP.parquet"),
        }
    if clean_file:
        inputs["clean"] = {
            "path": "inputs/no_need.xlsx",
            "original_name": _clean_original_name(clean_file[0], "no_need.xlsx"),
        }

    history_metadata: list[dict[str, str]] = []
    for index, (name, _source) in enumerate(history_items, start=1):
        suffix = Path(name).suffix.lower() or ".xlsx"
        if suffix not in {".xlsx", ".xlsm"}:
            suffix = ".xlsx"
        history_metadata.append(
            {
                "path": f"inputs/sales_history/sales_history_{index:03d}{suffix}",
                "original_name": _clean_original_name(name, f"sales_history_{index}{suffix}"),
            }
        )
    inputs["sales_history"] = history_metadata

    job: dict[str, object] = {
        "schema_version": SCHEMA_VERSION,
        "job_id": str(job_id).strip(),
        "created_at": datetime.now(timezone.utc).isoformat(),
        "requested_by": str(requested_by or "").strip(),
        "scope": normalized_scope,
        "salespeople": cleaned_salespeople,
        "batch_size": int(batch_size),
        "continue_after_failed_batch": bool(continue_after_failed_batch),
        "runtime_overrides": json_safe(dict(runtime_overrides)),
        "inputs": inputs,
    }

    destination = Path(output_path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(
        destination,
        mode="w",
        compression=zipfile.ZIP_DEFLATED,
        compresslevel=6,
        allowZip64=True,
    ) as zf:
        zf.writestr("job.json", json.dumps(job, indent=2, ensure_ascii=False))
        _write_stream(zf, str(inputs["stock"]["path"]), stock_file[1])  # type: ignore[index]
        _write_stream(zf, str(inputs["filters"]["path"]), filter_file[1])  # type: ignore[index]
        if profile_perm_file and isinstance(inputs["profile_perm"], dict):
            _write_stream(zf, str(inputs["profile_perm"]["path"]), profile_perm_file[1])
        if profile_temp_file and isinstance(inputs["profile_temp"], dict):
            _write_stream(zf, str(inputs["profile_temp"]["path"]), profile_temp_file[1])
        if clean_file and isinstance(inputs["clean"], dict):
            _write_stream(zf, str(inputs["clean"]["path"]), clean_file[1])
        for metadata, (_name, source) in zip(history_metadata, history_items):
            _write_stream(zf, metadata["path"], source)

    size = destination.stat().st_size
    if size >= MAX_RELEASE_ASSET_BYTES:
        destination.unlink(missing_ok=True)
        raise ValueError(
            "The uploaded files are too large for one full-processing request. "
            f"Bundle size: {size / (1024 ** 3):.2f} GiB."
        )
    return job
