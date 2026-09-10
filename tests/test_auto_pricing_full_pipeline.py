from __future__ import annotations

import io
import zipfile

import pytest

from adapters.auto_pricing_pipeline_adapter import (
    _extract_xlsx_zip,
    HISTORY_UPLOAD_NAMES,
    RAW_UPLOAD_NAMES,
    _named_xlsx_zip,
    default_auto_pricing_template_bytes,
    default_auto_pricing_template_path,
    default_remove_party_bytes,
    default_remove_party_path,
)


def test_git_default_removed_parties_is_available():
    assert default_remove_party_path().name == "remove_parties.xlsx"
    payload = default_remove_party_bytes()
    assert payload is not None
    assert payload[:2] == b"PK"


def test_uploaded_zip_extraction_is_flat_and_xlsx_only(tmp_path):
    payload = io.BytesIO()
    with zipfile.ZipFile(payload, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("nested/Cart.xlsx", b"cart")
        zf.writestr("nested/readme.txt", b"ignore")
    target = tmp_path / "out"
    _extract_xlsx_zip(payload.getvalue(), target, "test")
    assert (target / "Cart.xlsx").read_bytes() == b"cart"
    assert not (target / "readme.txt").exists()


def test_uploaded_zip_duplicate_workbook_names_are_rejected(tmp_path):
    payload = io.BytesIO()
    with zipfile.ZipFile(payload, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("a/Cart.xlsx", b"one")
        zf.writestr("b/Cart.xlsx", b"two")
    with pytest.raises(ValueError, match="duplicate workbook name"):
        _extract_xlsx_zip(payload.getvalue(), tmp_path / "out", "test")


def test_git_default_auto_pricing_template_is_available():
    assert default_auto_pricing_template_path().name == "auto_pricing_template.xlsx"
    payload = default_auto_pricing_template_bytes()
    assert payload is not None
    assert payload[:2] == b"PK"


def test_individual_uploads_are_packed_with_canonical_names():
    raw_files = {key: f"raw-{key}".encode() for key in RAW_UPLOAD_NAMES}
    payload = _named_xlsx_zip(raw_files, RAW_UPLOAD_NAMES, "raw")
    with zipfile.ZipFile(io.BytesIO(payload)) as zf:
        assert set(zf.namelist()) == set(RAW_UPLOAD_NAMES.values())
        assert zf.read("KSTAKA.xlsx") == b"raw-Stock"

    history_files = {key: f"history-{key}".encode() for key in HISTORY_UPLOAD_NAMES}
    payload = _named_xlsx_zip(history_files, HISTORY_UPLOAD_NAMES, "history")
    with zipfile.ZipFile(io.BytesIO(payload)) as zf:
        assert set(zf.namelist()) == set(HISTORY_UPLOAD_NAMES.values())
        assert zf.read("sales_120_file.xlsx") == b"history-Sales 120"
