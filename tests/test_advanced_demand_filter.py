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
