from __future__ import annotations

import io
import zipfile

import pytest
from openpyxl import load_workbook

from adapters.auto_pricing_pipeline_adapter import (
    ACTION_HEADERS,
    HISTORY_UPLOAD_NAMES,
    RAW_UPLOAD_NAMES,
    REQUIRED_HISTORY_UPLOADS,
    REQUIRED_RAW_UPLOADS,
    _extract_xlsx_zip,
    _named_xlsx_zip,
    default_remove_party_bytes,
    default_remove_party_path,
)


def test_saved_removed_parties_is_available():
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


def test_required_files_are_stock_sales_and_two_sales_history_files_only():
    assert REQUIRED_RAW_UPLOADS == {"Stock", "Sales"}
    assert REQUIRED_HISTORY_UPLOADS == {"Sales 120", "Sales 180"}
    assert "Stock" not in HISTORY_UPLOAD_NAMES


def test_optional_action_files_get_valid_empty_workbooks_when_omitted():
    raw_files = {"Stock": b"stock", "Sales": b"sales"}
    payload = _named_xlsx_zip(
        raw_files,
        RAW_UPLOAD_NAMES,
        "latest",
        required_keys=REQUIRED_RAW_UPLOADS,
        optional_headers=ACTION_HEADERS,
    )
    with zipfile.ZipFile(io.BytesIO(payload)) as zf:
        assert set(zf.namelist()) == set(RAW_UPLOAD_NAMES.values())
        assert zf.read("KSTAKA.xlsx") == b"stock"
        assert zf.read("SALE_DATA.xlsx") == b"sales"
        for dataset in ["Cart", "Hold", "Offer", "View", "Wishlist"]:
            data = zf.read(RAW_UPLOAD_NAMES[dataset])
            wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
            ws = wb[wb.sheetnames[0]]
            headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
            assert headers == ACTION_HEADERS[dataset]
            rows = list(ws.iter_rows(values_only=True))
            assert len(rows) == 1
            wb.close()


def test_optional_previous_action_files_get_valid_empty_workbooks_when_omitted():
    previous = {"Sales 120": b"s120", "Sales 180": b"s180"}
    payload = _named_xlsx_zip(
        previous,
        HISTORY_UPLOAD_NAMES,
        "previous",
        required_keys=REQUIRED_HISTORY_UPLOADS,
        optional_headers=ACTION_HEADERS,
    )
    with zipfile.ZipFile(io.BytesIO(payload)) as zf:
        assert zf.read("sales_120_file.xlsx") == b"s120"
        assert zf.read("sales_180_file.xlsx") == b"s180"
        for dataset in ["Cart", "Hold", "Offer", "View", "Wishlist"]:
            data = zf.read(HISTORY_UPLOAD_NAMES[dataset])
            wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
            ws = wb[wb.sheetnames[0]]
            headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
            assert headers == ACTION_HEADERS[dataset]
            rows = list(ws.iter_rows(values_only=True))
            assert len(rows) == 1
            wb.close()
