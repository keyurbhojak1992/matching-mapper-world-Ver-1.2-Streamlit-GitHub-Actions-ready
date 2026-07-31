from __future__ import annotations

import argparse
import json
import shutil
import stat
import sys
import traceback
import zipfile
from datetime import datetime, timezone
from pathlib import Path, PurePosixPath
from typing import Any

from adapters.suggested_adapter import run_suggested_batched_from_paths
from cloud_jobs.bundle import MAX_RELEASE_ASSET_BYTES, SCHEMA_VERSION

MAX_MEMBERS = 500
MAX_UNCOMPRESSED_BYTES = 10 * 1024 * 1024 * 1024
PROJECT_ROOT = Path(__file__).resolve().parents[1]
WORKER_VERSION = (PROJECT_ROOT / "VERSION.txt").read_text(encoding="utf-8").strip()

SET_RUNTIME_KEYS = {
    "SORT_HIGH_PRIORITY_PARTY_TYPES",
    "ALLOWED_MATCH_TYPES",
    "HIGH_PRIORITY_PARTY_TYPES",
    "GREEN_RGB_VALUES",
    "ORANGE_RGB_VALUES",
    "FILTER_RGB_VALUES",
}


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def _atomic_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    temporary.replace(path)


def _safe_member_path(root: Path, member_name: str) -> Path:
    pure = PurePosixPath(member_name)
    if pure.is_absolute() or ".." in pure.parts or not pure.parts:
        raise ValueError(f"Unsafe path in input bundle: {member_name!r}")
    destination = root.joinpath(*pure.parts).resolve()
    root_resolved = root.resolve()
    if destination != root_resolved and root_resolved not in destination.parents:
        raise ValueError(f"Input bundle path escapes extraction directory: {member_name!r}")
    return destination


def safe_extract_bundle(bundle_path: Path, destination: Path) -> None:
    """Extract an input ZIP while rejecting traversal, links and zip bombs."""
    destination.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(bundle_path) as archive:
        members = archive.infolist()
        if len(members) > MAX_MEMBERS:
            raise ValueError(f"Input bundle has too many members ({len(members)} > {MAX_MEMBERS}).")
        total = sum(info.file_size for info in members)
        if total > MAX_UNCOMPRESSED_BYTES:
            raise ValueError("Input bundle expands beyond the worker's safe disk limit.")

        for info in members:
            unix_mode = (info.external_attr >> 16) & 0xFFFF
            if stat.S_ISLNK(unix_mode):
                raise ValueError(f"Symbolic links are not allowed in input bundles: {info.filename}")
            target = _safe_member_path(destination, info.filename)
            if info.is_dir():
                target.mkdir(parents=True, exist_ok=True)
                continue
            target.parent.mkdir(parents=True, exist_ok=True)
            with archive.open(info, "r") as source, target.open("wb") as output:
                shutil.copyfileobj(source, output, length=1024 * 1024)


def _path_from_metadata(root: Path, metadata: object, *, required: bool, label: str) -> Path | None:
    if metadata is None:
        if required:
            raise ValueError(f"Job metadata is missing required input: {label}")
        return None
    if not isinstance(metadata, dict) or not metadata.get("path"):
        raise ValueError(f"Job metadata is invalid for input: {label}")
    path = _safe_member_path(root, str(metadata["path"]))
    if required and not path.is_file():
        raise FileNotFoundError(f"Required input is missing from the bundle: {label}")
    if not required and not path.is_file():
        return None
    return path


def _restore_runtime(payload: object) -> dict[str, object]:
    if not isinstance(payload, dict):
        raise ValueError("runtime_overrides must be a JSON object.")
    runtime: dict[str, object] = dict(payload)
    for key in SET_RUNTIME_KEYS:
        value = runtime.get(key)
        if isinstance(value, list):
            runtime[key] = set(value)
    return runtime


def _load_job(extracted_root: Path) -> dict[str, Any]:
    job_path = extracted_root / "job.json"
    if not job_path.is_file():
        raise FileNotFoundError("The input bundle does not contain job.json.")
    job = json.loads(job_path.read_text(encoding="utf-8"))
    if int(job.get("schema_version", -1)) != SCHEMA_VERSION:
        raise ValueError(
            f"Unsupported job schema {job.get('schema_version')!r}; expected {SCHEMA_VERSION}."
        )
    return job


def run_job(arguments: argparse.Namespace) -> int:
    status_path = Path(arguments.status)
    failure_path = Path(arguments.failure)
    result_path = Path(arguments.result)
    manifest_path = Path(arguments.manifest)
    processing_log_path = Path(arguments.log)
    work_dir = Path(arguments.work_dir)
    extracted_root = work_dir / "input"

    status_base: dict[str, Any] = {
        "schema_version": 2,
        "worker_version": WORKER_VERSION,
        "state": "starting",
        "message": "Preparing full processing.",
        "started_at": _utc_now(),
        "updated_at": _utc_now(),
    }
    _atomic_json(status_path, status_base)

    try:
        if work_dir.exists():
            shutil.rmtree(work_dir)
        extracted_root.mkdir(parents=True, exist_ok=True)
        safe_extract_bundle(Path(arguments.bundle), extracted_root)
        job = _load_job(extracted_root)

        inputs = job.get("inputs")
        if not isinstance(inputs, dict):
            raise ValueError("Job metadata does not contain a valid inputs object.")
        stock_path = _path_from_metadata(extracted_root, inputs.get("stock"), required=True, label="stock")
        filter_path = _path_from_metadata(extracted_root, inputs.get("filters"), required=False, label="filters")
        profile_perm_path = _path_from_metadata(
            extracted_root, inputs.get("profile_perm"), required=False, label="permanent profile"
        )
        profile_temp_path = _path_from_metadata(
            extracted_root, inputs.get("profile_temp"), required=False, label="temporary profile"
        )
        clean_path = _path_from_metadata(extracted_root, inputs.get("clean"), required=False, label="clean file")
        history_metadata = inputs.get("sales_history") or []
        if not isinstance(history_metadata, list):
            raise ValueError("sales_history metadata must be a list.")
        history_paths = [
            path
            for item in history_metadata
            if (path := _path_from_metadata(extracted_root, item, required=False, label="sales history"))
            is not None
        ]

        runtime = _restore_runtime(job.get("runtime_overrides"))
        scope = str(job.get("scope", "selected"))
        salespeople = job.get("salespeople") or []
        if not isinstance(salespeople, list):
            raise ValueError("salespeople must be a list.")
        batch_size = int(job.get("batch_size", 3))
        continue_after_failure = bool(job.get("continue_after_failed_batch", True))
        job_id = str(job.get("job_id", ""))

        status_base.update(
            {
                "job_id": job_id,
                "requested_by": str(job.get("requested_by", "")),
                "scope": scope,
                "batch_size": batch_size,
                "state": "running",
                "message": "Files are ready; starting processing.",
                "updated_at": _utc_now(),
            }
        )
        _atomic_json(status_path, status_base)

        latest_progress: dict[str, object] = {}

        def update_status(progress: dict[str, object]) -> None:
            latest_progress.clear()
            latest_progress.update(progress)
            payload = dict(status_base)
            payload.update(latest_progress)
            payload["job_id"] = job_id
            payload["updated_at"] = _utc_now()
            _atomic_json(status_path, payload)

        assert stock_path is not None
        assert filter_path is not None
        manifest = run_suggested_batched_from_paths(
            stock_path=stock_path,
            profile_perm_path=profile_perm_path,
            profile_temp_path=profile_temp_path,
            filter_path=filter_path,
            clean_path=clean_path,
            sales_history_paths=history_paths,
            runtime_overrides=runtime,
            scope=scope,
            salespeople=[str(item) for item in salespeople],
            batch_size=batch_size,
            output_path=result_path,
            continue_after_failed_batch=continue_after_failure,
            status_callback=update_status,
            manifest_path=manifest_path,
            processing_log_path=processing_log_path,
        )

        if result_path.is_file() and result_path.stat().st_size >= MAX_RELEASE_ASSET_BYTES:
            size_gib = result_path.stat().st_size / (1024 ** 3)
            result_path.unlink(missing_ok=True)
            raise ValueError(
                "The combined result exceeds the maximum supported download size "
                f"({size_gib:.2f} GiB). Reduce output options or submit a smaller salesperson selection."
            )

        final_state = str(manifest.get("final_state", "failed"))
        total_batches = int(manifest.get("total_batches", 0) or 0)
        final_status = dict(status_base)
        final_status.update(latest_progress)
        final_status.update(
            {
                "state": final_state,
                "stage": "completed",
                "message": (
                    "All batches completed and the final combined package is ready."
                    if final_state == "completed"
                    else "Final combined packaging finished with one or more failed or skipped batches."
                ),
                # Keep the final progress values instead of dropping them when
                # the worker writes its last status payload. This prevents the
                # completed dashboard from falling back to "Waiting".
                "current_batch": total_batches,
                "total_salespeople": manifest.get("total_salespeople", 0),
                "total_batches": total_batches,
                "successful_batches": manifest.get("summary", {}).get("successful", 0),
                "failed_batches": manifest.get("summary", {}).get("failed", 0),
                "skipped_batches": manifest.get("summary", {}).get("skipped", 0),
                "result_filename": result_path.name,
                "output_layout": manifest.get("output_layout", {}),
                "completed_at": manifest.get("completed_at", _utc_now()),
                "updated_at": _utc_now(),
            }
        )
        _atomic_json(status_path, final_status)
        return 0 if final_state in {"completed", "partial_success"} else 1
    except Exception as exc:
        failure_text = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
        failure_path.parent.mkdir(parents=True, exist_ok=True)
        failure_path.write_text(failure_text, encoding="utf-8")
        failed_status = dict(status_base)
        failed_status.update(
            {
                "state": "failed",
                "message": str(exc),
                "completed_at": _utc_now(),
                "updated_at": _utc_now(),
            }
        )
        _atomic_json(status_path, failed_status)
        print(failure_text, file=sys.stderr)
        return 1


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Run one Matching Mapper job bundle")
    parser.add_argument("--bundle", type=Path, required=True)
    parser.add_argument("--work-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--status", type=Path, required=True)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--log", type=Path, required=True)
    parser.add_argument("--failure", type=Path, required=True)
    return parser


def main() -> int:
    return run_job(_parser().parse_args())


if __name__ == "__main__":
    raise SystemExit(main())
