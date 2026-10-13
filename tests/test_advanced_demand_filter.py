from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from adapters.advanced_demand_filter_adapter import run_advanced_demand_filter
from adapters.common import MASTER_CODE_DIR, load_fresh_module


def _stock_bytes() -> bytes:
    stock = pd.DataFrame([
        {
            "Serial No": "1001",
            "Shp": "RBC",
            "Carat": 1.05,
            "Color": "D",
            "Qua": "VS1",
            "Flo": "NON",
            "Lab": "GIA",
            "Cut": "EX",
            "Pol": "EX",
            "Sym": "EX",
            "Shade": "NONE",
            "Milky": "NONE",
            "Disc%": -35,
            "Rate": 5000,
            "Amount": 5250,
            "SGPComment": "",
        }
    ])
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        stock.to_excel(writer, index=False, sheet_name="Sheet1")
    return buffer.getvalue()


def _filter_bytes() -> bytes:
    rules = pd.DataFrame([
        {
            "Sold Party": "TEST PARTY",
            "Sales_Person": "Mr. Test User",
            "Remarks": "INCLUDE",
            "Shp": "RBC",
            "Carat": "1.00 TO 1.10",
            "COLOR NAME": "D",
            "Qua": "VS1",
            "Flo": "NON",
            "Lab": "GIA",
        }
    ])
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        pd.DataFrame(columns=["Sold Party"]).to_excel(
            writer, index=False, sheet_name="Removed Party"
        )
        rules.to_excel(writer, index=False, sheet_name="Change Format")
    return buffer.getvalue()


def test_advanced_demand_defaults_match_requested_configuration():
    module = load_fresh_module(
        MASTER_CODE_DIR / "Advance_Demand_Filter_Process.py",
        "test_advanced_demand_defaults",
    )
    assert module.TOLERANCE_LEVEL == 1
    assert module.ENABLE_FANCY_COLOR_TOLERANCE is False
    assert module.ENABLE_FLO_MAPPING is False
    assert module.ENABLE_RELATED_TAG is False
    assert module.ENABLE_CLEAN_FILE is True
    assert module.MATCH_STATUS_FILTER_MODE == "NONE"
    assert module.MATCH_STATUS_ALLOWED_TYPES == ["Exact", "Near by"]
    assert module.MATCH_STATUS_SPECIFIC_PARTIES == ["NONE"]


def test_advanced_demand_runs_without_optional_clean_file():
    output, filename, logs = run_advanced_demand_filter(
        stock_bytes=_stock_bytes(),
        filter_bytes=_filter_bytes(),
        clean_bytes=None,
    )
    assert filename.endswith(".zip")
    assert "No Need workbook not supplied" in logs
    with zipfile.ZipFile(io.BytesIO(output)) as archive:
        assert archive.namelist() == ["Mr Test User.xlsx"]
        workbook = load_workbook(io.BytesIO(archive.read("Mr Test User.xlsx")), data_only=True)
        assert workbook.sheetnames == ["Navigation", "TEST PARTY"]
        assert workbook["TEST PARTY"].max_row == 2


def test_existing_demand_mapper_is_still_separate():
    root = Path(__file__).resolve().parents[1]
    app_text = (root / "app.py").read_text(encoding="utf-8")
    assert '"Demand Filter Mapper"' in app_text
    assert '"Advanced Demand Filter"' in app_text
    assert "run_filter_demand(" in app_text
    assert "run_advanced_demand_filter(" in app_text


def _multi_party_stock_and_filters() -> tuple[bytes, bytes]:
    """Four distinguishable stones, four party scopes, and two salespeople."""
    stock = pd.DataFrame([
        {
            "Serial No": f"100{i}", "Shp": "RBC", "Carat": 1.05,
            "Color": color, "Qua": "VS1", "Flo": "NON", "Lab": "GIA",
            "Cut": "EX", "Pol": "EX", "Sym": "EX", "Shade": "NONE",
            "Milky": "NONE", "Disc%": -35, "Rate": 5000,
            "Amount": 5250, "SGPComment": "",
        }
        for i, color in enumerate(["D", "E", "F", "G"], start=1)
    ])
    rules = []
    scopes = {
        ("Mr. First Seller", "SHORT PARTY"): "D",
        ("Mr. First Seller", "TWO STONES"): "DE",
        ("Mr. First Seller", "THREE STONES"): "DEF",
        ("Mr. Second Seller", "ONLY ONE"): "D",
    }
    for (salesperson, party), colors in scopes.items():
        for color in colors:
            rules.append({
                "Sold Party": party, "Sales_Person": salesperson,
                "Remarks": "INCLUDE", "Shp": "RBC",
                "Carat": "1.00 TO 1.10", "COLOR NAME": color,
                "Qua": "VS1", "Flo": "NON", "Lab": "GIA",
            })
    stock_buffer, rules_buffer = io.BytesIO(), io.BytesIO()
    with pd.ExcelWriter(stock_buffer, engine="openpyxl") as writer:
        stock.to_excel(writer, sheet_name="Sheet1", index=False)
    with pd.ExcelWriter(rules_buffer, engine="openpyxl") as writer:
        pd.DataFrame(columns=["Sold Party"]).to_excel(writer, sheet_name="Removed Party", index=False)
        pd.DataFrame(rules).to_excel(writer, sheet_name="Change Format", index=False)
    return stock_buffer.getvalue(), rules_buffer.getvalue()


def _run_layout(layout, threshold):
    stock, filters = _multi_party_stock_and_filters()
    return run_advanced_demand_filter(
        stock_bytes=stock,
        filter_bytes=filters,
        min_stones_per_party=threshold,
        output_layout=layout,
        runtime_overrides={"TOLERANCE_LEVEL": 0},
    )


def test_threshold_removes_party_sheet_navigation_row_and_empty_salesperson():
    output, _name, logs = _run_layout("salesperson_workbooks", 2)
    assert "Party sheets removed below minimum: 2" in logs
    with zipfile.ZipFile(io.BytesIO(output)) as outer:
        assert outer.namelist() == ["Mr First Seller.xlsx"]
        wb = load_workbook(io.BytesIO(outer.read("Mr First Seller.xlsx")))
        assert wb.sheetnames == ["Navigation", "TWO STONES", "THREE STONES"]
        nav = wb["Navigation"]
        assert nav.max_row == 3
        assert [(nav.cell(r, 2).value, nav.cell(r, 3).value) for r in (2, 3)] == [
            ("TWO STONES", 2), ("THREE STONES", 3)
        ]
        for row in (2, 3):
            target = nav.cell(row, 6).hyperlink.target
            assert target == f"#'{nav.cell(row, 1).value}'!A1"
            assert nav.cell(row, 1).value in wb.sheetnames
        assert [wb[title].max_row - 1 for title in wb.sheetnames[1:]] == [2, 3]
        wb.close()


def test_equal_to_threshold_is_retained_and_above_comes_through():
    output, _name, _logs = _run_layout("salesperson_workbooks", 3)
    with zipfile.ZipFile(io.BytesIO(output)) as outer:
        wb = load_workbook(io.BytesIO(outer.read("Mr First Seller.xlsx")))
        assert wb.sheetnames == ["Navigation", "THREE STONES"]
        assert wb["Navigation"]["C2"].value == 3
        wb.close()


def test_separate_layout_nested_zip_contains_only_retained_party_workbooks():
    output, _name, logs = _run_layout("separate_party_files", 2)
    assert "Separated party Excel files created: 2" in logs
    with zipfile.ZipFile(io.BytesIO(output)) as outer:
        assert outer.namelist() == ["Mr First Seller.zip"]
        with zipfile.ZipFile(io.BytesIO(outer.read("Mr First Seller.zip"))) as inner:
            assert inner.namelist() == ["TWO STONES.xlsx", "THREE STONES.xlsx"]
            for party_name, count in (("TWO STONES", 2), ("THREE STONES", 3)):
                wb = load_workbook(io.BytesIO(inner.read(party_name + ".xlsx")))
                assert wb.sheetnames == [party_name]
                assert wb[party_name].max_row - 1 == count
                assert wb[party_name]["A2"].value == 1
                assert "Match_Type" not in [c.value for c in wb[party_name][1]]
                wb.close()


def test_both_layout_preserves_regular_and_party_content():
    output, _name, _logs = _run_layout("both", 2)
    with zipfile.ZipFile(io.BytesIO(output)) as outer:
        assert outer.namelist() == ["Mr First Seller.xlsx", "Mr First Seller.zip"]
        regular = load_workbook(io.BytesIO(outer.read("Mr First Seller.xlsx")))
        with zipfile.ZipFile(io.BytesIO(outer.read("Mr First Seller.zip"))) as inner:
            for party in ("TWO STONES", "THREE STONES"):
                separate = load_workbook(io.BytesIO(inner.read(f"{party}.xlsx")))
                assert list(regular[party].values) == list(separate[party].values)
                from copy import copy
                assert copy(regular[party]["A1"].font) == copy(separate[party]["A1"].font)
                assert copy(regular[party]["A1"].fill) == copy(separate[party]["A1"].fill)
                separate.close()
        regular.close()


def test_empty_result_zip_when_all_parties_below_minimum():
    output, _name, logs = _run_layout("both", 4)
    with zipfile.ZipFile(io.BytesIO(output)) as outer:
        assert outer.namelist() == []
    assert "output ZIP is empty" in logs


def test_invalid_advanced_demand_output_controls_are_rejected():
    import pytest
    stock, filters = _multi_party_stock_and_filters()
    with pytest.raises(ValueError, match="Minimum stones"):
        run_advanced_demand_filter(stock_bytes=stock, filter_bytes=filters, min_stones_per_party=-1)
    with pytest.raises(ValueError, match="Unknown Advanced Demand output"):
        run_advanced_demand_filter(stock_bytes=stock, filter_bytes=filters, output_layout="unexpected")


def test_make_config_controls_exposed_without_changing_standard_mapper():
    content = (Path(__file__).resolve().parents[1] / "app.py").read_text(encoding="utf-8")
    advanced_page = content.split("def advanced_demand_filter_page() -> None:")[1].split("def sorting_page() -> None:")[0]
    assert 'st.popover("⚙️ Make Config"' in advanced_page
    assert "advanced_demand_min_stones_per_party" in advanced_page
    assert "advanced_demand_output_layout" in advanced_page
    assert "run_filter_demand(" not in advanced_page


def test_streamed_large_report_matches_standard_content_style_and_hyperlinks(monkeypatch):
    """Protect existing Excel layout if a large report selects write-only mode."""
    import processors.advanced_demand_batch as batch
    stock, rules = _multi_party_stock_and_filters()
    sheet = pd.read_excel(io.BytesIO(stock), sheet_name="Sheet1")
    sheet["CERTIFICATE"] = [f"https://example.test/cert/{i}" for i in range(len(sheet))]
    stock_buffer = io.BytesIO()
    with pd.ExcelWriter(stock_buffer, engine="openpyxl") as writer:
        sheet.to_excel(writer, index=False, sheet_name="Sheet1")
    stock = stock_buffer.getvalue()

    monkeypatch.setattr(batch, "STREAM_REPORT_ROW_THRESHOLD", 1_000_000)
    normal, _, _ = run_advanced_demand_filter(
        stock_bytes=stock, filter_bytes=rules, output_layout="both"
    )
    monkeypatch.setattr(batch, "STREAM_REPORT_ROW_THRESHOLD", 1)
    streamed, _, _ = run_advanced_demand_filter(
        stock_bytes=stock, filter_bytes=rules, output_layout="both"
    )

    def get_xlsx_members(payload):
        with zipfile.ZipFile(io.BytesIO(payload)) as outer:
            result = {}
            for file_name in outer.namelist():
                if file_name.endswith(".zip"):
                    with zipfile.ZipFile(io.BytesIO(outer.read(file_name))) as nested:
                        for nested_name in nested.namelist():
                            result[(file_name, nested_name)] = nested.read(nested_name)
                else:
                    result[(file_name,)] = outer.read(file_name)
            return result

    normal_files = get_xlsx_members(normal)
    streamed_files = get_xlsx_members(streamed)
    assert normal_files.keys() == streamed_files.keys()
    for name in normal_files:
        normal_book = load_workbook(io.BytesIO(normal_files[name]))
        streamed_book = load_workbook(io.BytesIO(streamed_files[name]))
        try:
            assert normal_book.sheetnames == streamed_book.sheetnames
            for sheet_name in normal_book.sheetnames:
                original = normal_book[sheet_name]
                optimized = streamed_book[sheet_name]
                assert original.max_row == optimized.max_row
                assert original.max_column == optimized.max_column
                assert list(original.values) == list(optimized.values)
                for row in original.iter_rows():
                    for orig_cell in row:
                        actual = optimized[orig_cell.coordinate]
                        assert (
                            actual.hyperlink.target if actual.hyperlink else None
                        ) == (orig_cell.hyperlink.target if orig_cell.hyperlink else None)
                        for prop in ("font", "fill", "border", "alignment", "number_format"):
                            assert str(getattr(actual, prop)) == str(getattr(orig_cell, prop))
        finally:
            normal_book.close()
            streamed_book.close()


def test_predicate_cache_does_not_confuse_include_exclude_or_party_metadata():
    """A cache hit can reuse matches, never party metadata, tags or rule remarks."""
    stock = _stock_bytes()
    filters = pd.DataFrame([
        {"Sold Party": "ONE", "Sales_Person": "Agent", "Country": "IN", "Remarks": "INCLUDE", "Shp": "RBC"},
        {"Sold Party": "ONE", "Sales_Person": "Agent", "Country": "US", "Remarks": "INCLUDE", "Shp": "RBC"},
        {"Sold Party": "TWO", "Sales_Person": "Agent", "Country": "UK", "Remarks": "INCLUDE", "Shp": "RBC"},
        {"Sold Party": "THREE", "Sales_Person": "Agent", "Country": "FR", "Remarks": "EXCLUDE", "Shp": "RBC"},
    ])
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        filters.to_excel(writer, sheet_name="Change Format", index=False)
    output, _, logs = run_advanced_demand_filter(stock_bytes=stock, filter_bytes=buf.getvalue())
    assert "Unique filter rules processed: 4" in logs
    assert "Matching predicate results reused: 2" in logs
    with zipfile.ZipFile(io.BytesIO(output)) as archive:
        workbook = load_workbook(io.BytesIO(archive.read("Agent.xlsx")))
        try:
            assert workbook.sheetnames == ["Navigation", "ONE", "TWO"]
            assert workbook["Navigation"]["D2"].value == "IN"
            assert workbook["Navigation"]["D3"].value == "UK"
        finally:
            workbook.close()
