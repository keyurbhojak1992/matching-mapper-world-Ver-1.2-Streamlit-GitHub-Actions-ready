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


def test_auto_pricing_builder_tolerates_trailing_blank_action_cells(tmp_path):
    """A blank trailing Sold Party cell must not make an action row unreadable."""
    from datetime import datetime
    from openpyxl import Workbook
    from processors.auto_pricing_finalizer_v4_1_3.build_auto_pricing_v4_1_3 import build_maps

    def save_book(name, headers, rows):
        wb = Workbook()
        ws = wb.active
        ws.append(headers)
        for row in rows:
            ws.append(row)
        wb.save(tmp_path / name)

    sales_headers = [
        'Serial No', 'Shp', 'Carat', 'Carat_Grp', 'Col', 'Qua', 'Cut', 'Pol', 'Sym',
        'Lab', 'Flo', 'Sold Party', 'Sale Date', 'SPer', 'SRate', 'Amt'
    ]
    sale_row = [
        1, 'RND', 1.0, '1.00-1.49', 'D', 'VS1', 'EX', 'EX', 'EX',
        'GIA', 'NON', 'Buyer', datetime(2026, 8, 20), 0, 100, 100
    ]
    save_book('sales_120_file.xlsx', sales_headers, [sale_row])
    save_book('sales_180_file.xlsx', sales_headers, [sale_row])

    # Excel omits the final blank cell from the stored row. This used to trigger
    # IndexError when the builder accessed the Sold Party column by position.
    save_book(
        'Hold.xlsx',
        ['SERIAL_NO', 'JANGAD_DATE', 'Date', 'PARTY_CODE', 'Sold Party'],
        [[123, datetime(2026, 8, 20), datetime(2026, 8, 20), 'P001']],
    )
    save_book('Wishlist.xlsx', ['CREATED_DATETIME', 'Sold Party', 'Serial No'], [])
    save_book('View.xlsx', ['SEQ_NO', 'Sold Party', 'SERIAL_NO'], [])
    save_book('Offer.xlsx', ['CREATED_DATETIME', 'Sold Party', 'Serial No'], [])
    save_book('Cart.xlsx', ['CREATED_DATETIME', 'Sold Party', 'Serial No'], [])

    avg_rate, sale_count, buyers, action_parties = build_maps(tmp_path)
    assert avg_rate
    assert sale_count
    assert buyers
    assert dict(action_parties) == {}


def test_exact_output_duplicate_cleaning_keeps_nonidentical_activity():
    from datetime import datetime
    from processors.auto_pricing_finalizer_v4_1_3.process_datasets_v4_1_3 import remove_exact_output_duplicates

    headers = ["SERIAL_NO", "JANGAD_DATE", "PARTY_CODE", "Sold Party"]
    base = {"SERIAL_NO": "412429", "JANGAD_DATE": datetime(2026, 8, 21, 0, 7, 11), "PARTY_CODE": "P1", "Sold Party": "BUYER A"}
    same = dict(base)
    different = dict(base)
    different["PARTY_CODE"] = "P2"
    kept, removed = remove_exact_output_duplicates([base, same, different], headers)
    assert removed == 1
    assert len(kept) == 2
    assert kept[0]["PARTY_CODE"] == "P1"
    assert kept[1]["PARTY_CODE"] == "P2"
