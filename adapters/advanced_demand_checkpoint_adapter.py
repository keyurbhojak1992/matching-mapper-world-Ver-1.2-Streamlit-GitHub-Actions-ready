"""File-backed, resumable Streamlit job controller for Advanced Demand Filter only.

Jobs can resume between Streamlit reruns or browser reconnects while the same
server's temporary files remain available. It is NOT durable cloud storage.
"""
from __future__ import annotations

from datetime import datetime
import json
import os
from pathlib import Path
import secrets
import shutil
import sys
import gc
import tempfile

try:
    import fcntl  # POSIX Streamlit hosts; locks are released on worker termination.
except ImportError:  # pragma: no cover - Windows local development
    fcntl = None
import time

from openpyxl import load_workbook

from processors.advanced_demand_incremental import get_state, run_checkpoint_step
from .advanced_demand_filter_adapter import MASTER_PATH
from .common import apply_runtime_overrides, load_fresh_module

JOBS_ROOT = Path(tempfile.gettempdir()) / "matching_mapper_advanced_demand_jobs"
MAX_JOB_AGE_SECONDS = 24 * 3600


def _job_folder(job_id: str) -> Path:
    # Prevent relative paths and directory traversal in pasted resume tokens.
    if len(job_id) != 40 or any(char not in "0123456789abcdef" for char in job_id):
        raise ValueError("Invalid batch job ID")
    return JOBS_ROOT / job_id


def cleanup_old_jobs():
    JOBS_ROOT.mkdir(parents=True, exist_ok=True)
    now = time.time()
    for folder in JOBS_ROOT.iterdir():
        if not folder.is_dir():
            continue
        try:
            if now - folder.stat().st_mtime > MAX_JOB_AGE_SECONDS:
                shutil.rmtree(folder)
        except OSError:
            pass


def create_batch_job(*, stock_bytes: bytes, filter_bytes: bytes,
                     clean_bytes: bytes | None = None,
                     runtime_overrides: dict | None = None,
                     batch_rows: int = 1000,
                     min_stones_per_party: int = 0,
                     output_layout: str = "salesperson_workbooks") -> str:
    cleanup_old_jobs()
    job_id = secrets.token_hex(20)
    directory = _job_folder(job_id)
    directory.mkdir(mode=0o700)
    (directory / "stock.xlsx").write_bytes(stock_bytes)
    (directory / "filters.xlsx").write_bytes(filter_bytes)
    if clean_bytes:
        (directory / "no_need.xlsx").write_bytes(clean_bytes)
    try:
        filter_wb = load_workbook(directory / "filters.xlsx", read_only=True)
        try:
            estimated_rows = max(0, (filter_wb["Change Format"].max_row or 1) - 1)
        finally:
            filter_wb.close()
    except Exception:
        estimated_rows = 0
    meta = {
        "runtime_overrides": runtime_overrides or {},
        "batch_rows": int(batch_rows),
        "min_stones_per_party": int(min_stones_per_party),
        "output_layout": output_layout,
        "estimated_rows": estimated_rows,
        "output_name": f"Advanced_Demand_Filter_Matches_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.zip",
    }
    (directory / "config.json").write_text(json.dumps(meta), encoding="utf-8")
    return job_id


def batch_job_snapshot(job_id: str) -> tuple[dict, dict]:
    folder = _job_folder(job_id)
    config = folder / "config.json"
    if not config.is_file():
        raise FileNotFoundError("This job is not available on the current server. Upload the files to start again.")
    meta = json.loads(config.read_text(encoding="utf-8"))
    state = get_state(folder) or {"phase": "ready", "filter_offset": 0, "report_index": 0,
                                   "processed_filter_rows": 0, "reports_created": 0}
    return meta, state


def process_next_batch(job_id: str) -> dict:
    folder = _job_folder(job_id)
    # Prevent two open tabs from processing the same SQLite cursor simultaneously.
    with (folder / ".job.lock").open("a+b") as lock:
        if fcntl is not None:
            fcntl.flock(lock.fileno(), fcntl.LOCK_EX)
        try:
            return _process_next_batch_locked(job_id, folder)
        finally:
            if fcntl is not None:
                fcntl.flock(lock.fileno(), fcntl.LOCK_UN)


def _process_next_batch_locked(job_id: str, folder: Path) -> dict:
    config, _ = batch_job_snapshot(job_id)
    module = load_fresh_module(MASTER_PATH, "advanced_demand_filter_batch_master")
    apply_runtime_overrides(module, config["runtime_overrides"])
    module.FLUORESCENCE_ORDER = (
        ["NON", "FNT", "MED", "STG", "VST"]
        if bool(getattr(module, "ENABLE_FLO_MAPPING", False))
        else ["NON", "VSL", "SL", "FNT", "MED", "STG", "VST"]
    )
    if hasattr(module, "ORDER_MAPPING"):
        module.ORDER_MAPPING["FLO"] = module.FLUORESCENCE_ORDER
    print(f"[Advanced Demand batch {job_id[:8]}] starting next checkpoint", flush=True)
    try:
        state = run_checkpoint_step(
            module,
        str(folder / "stock.xlsx"),
        str(folder / config["output_name"]),
        str(folder / "filters.xlsx"),
        int(getattr(module, "TOLERANCE_LEVEL", 1)),
        str(folder / "no_need.xlsx") if (folder / "no_need.xlsx").is_file() else None,
        job_dir=folder,
        batch_rows=config["batch_rows"],
        max_salespeople=1,
        min_stones_per_party=config["min_stones_per_party"],
        output_layout=config["output_layout"],
        )
    finally:
        # load_fresh_module registers a new name in sys.modules on every call.
        # Remove it after each checkpoint to avoid memory growth across 200+ reruns.
        sys.modules.pop(module.__name__, None)
        del module
        gc.collect()
    os.utime(folder, None)
    print(f"[Advanced Demand batch {job_id[:8]}] phase={state['phase']} "
          f"source_rows={state['filter_offset']:,} "
          f"salespeople={state['report_index']:,}", flush=True)
    return state


def batch_output_path(job_id: str) -> Path:
    folder = _job_folder(job_id)
    meta, state = batch_job_snapshot(job_id)
    if state["phase"] != "done":
        raise RuntimeError("Batch report is not complete")
    path = folder / meta["output_name"]
    if not path.is_file():
        raise FileNotFoundError("Completed report ZIP is missing from temporary storage")
    return path
