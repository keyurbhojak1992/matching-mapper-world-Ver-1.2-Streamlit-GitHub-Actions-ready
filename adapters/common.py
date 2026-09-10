from __future__ import annotations

import contextlib
import hashlib
import importlib.util
import io
import os
import re
import sys
import tempfile
import uuid
from pathlib import Path
from types import ModuleType
from typing import BinaryIO, Iterable


ROOT_DIR = Path(__file__).resolve().parents[1]
MASTER_CODE_DIR = ROOT_DIR / "original_master_code"
DATA_DIR = ROOT_DIR / "data"

DEFAULT_DATA_FILES = {
    "profile_perm": "PARTY_PROFILE_PERM.parquet",
    "profile_temp": "PARTY_PROFILE_TEMP.parquet",
    "filters": "filters - EBID.xlsx",
    "no_need": "no_need.xlsx",
    "auto_pricing_remove_parties": "remove_parties.xlsx",
}


def repository_data_path(key: str) -> Path:
    """Return the exact repository data path configured for an optional upload."""
    try:
        filename = DEFAULT_DATA_FILES[key]
    except KeyError as exc:
        raise KeyError(f"Unknown repository data key: {key}") from exc
    return DATA_DIR / filename


def resolve_repository_input(
    supplied_path: Path | None,
    key: str,
    *,
    required: bool = False,
) -> Path | None:
    """Prefer an uploaded path, otherwise use the exact file under ``data``."""
    if supplied_path is not None:
        candidate = Path(supplied_path)
        if candidate.is_file():
            return candidate
    fallback = repository_data_path(key)
    if fallback.is_file():
        return fallback
    if required:
        raise FileNotFoundError(
            f"No uploaded file was provided and the repository default is missing: data/{fallback.name}"
        )
    return None


def load_fresh_module(path: Path, prefix: str) -> ModuleType:
    """Load a Python source file under a unique module name for session isolation."""
    module_name = f"{prefix}_{uuid.uuid4().hex}"
    spec = importlib.util.spec_from_file_location(module_name, path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Unable to load master module: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    try:
        spec.loader.exec_module(module)
    except Exception:
        sys.modules.pop(module_name, None)
        raise
    return module


def apply_runtime_overrides(module: ModuleType, overrides: dict[str, object] | None) -> None:
    if not overrides:
        return
    for name, value in overrides.items():
        if hasattr(module, name):
            setattr(module, name, value)


def write_bytes(path: Path, data: bytes | None) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data or b"")
    return path


def safe_filename(name: str, fallback: str = "output", max_length: int = 120) -> str:
    cleaned = re.sub(r'[\\/*?:"<>|]+', "_", str(name or "")).strip().rstrip(". ")
    return (cleaned or fallback)[:max_length]


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


@contextlib.contextmanager
def captured_output():
    stream = io.StringIO()
    with contextlib.redirect_stdout(stream), contextlib.redirect_stderr(stream):
        yield stream


@contextlib.contextmanager
def temp_workspace(prefix: str):
    with tempfile.TemporaryDirectory(prefix=prefix) as tmp:
        yield Path(tmp)
