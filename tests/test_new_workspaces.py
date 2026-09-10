from __future__ import annotations

import hashlib
import io
import subprocess
import sys
import zipfile
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
from openpyxl import Workbook, load_workbook

from adapters.auto_pricing_adapter import AUTOPRICE_VERSION, PACKAGE_DIR as AUTOPRICE_PACKAGE_DIR, run_auto_pricing
from adapters.dealmaker_adapter import DEFAULT_CONTROL, PACKAGE_DIR, RUNNER, bundled_dealmaker_control_bytes
from adapters.layout_stock_matcher_adapter import run_layout_stock_matcher

ROOT = Path(__file__).resolve().parents[1]


def _xlsx_bytes(frame: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    frame.to_excel(buffer, index=False)
    return buffer.getvalue()


def _zip_bytes(files: dict[str, bytes]) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, payload in files.items():
            zf.writestr(name, payload)
    return buffer.getvalue()


def _auto_pricing_template_bytes() -> bytes:
    cart_headers = ["SERIAL_NO", "CREATED_DATETIME", "Sold Party"]
    hold_headers = ["SERIAL_NO", "JANGAD_DATE", "Sold Party"]
    view_headers = ["SEQ_NO", "SERIAL_NO", "CREATED_DATETIME", "Sold Party"]
    sheets = {
        "Sheet1": [
            "Serial No", "Shp", "Carat", "Carat_Grp", "Color name", "Qua", "Cut", "Pol", "Sym",
            "Lab", "Flo", "GRap", "GPer", "Expected", "Sold_Avg_Rate_120days", "Diff", "Remark",
            "Tag", "Unique_Buyer_Count", "Sold_Party", "sale_count", "Action_count", "ALL_Action_count",
        ],
        "Sale_segment": [
            "Serial No", "Shp", "Carat", "Carat_Grp", "Col", "Qua", "Cut", "Pol", "Sym", "Lab", "Flo",
            "Sold Party", "Sale Date", "SPer", "SRate", "Amt", "Date", "Tag",
        ],
        "Avg_Rate": [
            "Serial No", "Shp", "Carat", "Carat_Grp", "Col", "Qua", "Cut", "Pol", "Sym", "Lab", "Flo",
            "Sold Party", "Sale Date", "SPer", "SRate", "Amt", "Date", "Tag",
        ],
        "Hold": hold_headers + ["Date"],
        "Wishlist": cart_headers,
        "View": view_headers + ["Date"],
        "Offer": cart_headers,
        "Cart": cart_headers,
    }
    wb = Workbook()
    wb.remove(wb.active)
    for sheet_name, headers in sheets.items():
        ws = wb.create_sheet(sheet_name)
        ws.append(headers)
        ws.append([None] * len(headers))
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def test_auto_pricing_v4_1_3_package_is_bundled():
    assert AUTOPRICE_VERSION == "4.1.3"
    for filename in [
        "process_datasets_v4_1_3.py",
        "build_auto_pricing_v4_1_3.py",
        "run_all_v4_1_3.py",
        "column_mapping_aliases.json",
    ]:
        assert (AUTOPRICE_PACKAGE_DIR / filename).is_file()


def test_auto_pricing_adapter_runs_v4_1_3_end_to_end():
    today = datetime.now(ZoneInfo("Asia/Kolkata")).date().isoformat()
    cart_headers = ["SERIAL_NO", "CREATED_DATETIME", "Sold Party"]
    hold_headers = ["SERIAL_NO", "JANGAD_DATE", "Sold Party"]
    view_headers = ["SEQ_NO", "SERIAL_NO", "CREATED_DATETIME", "Sold Party"]
    sales_headers = ["Serial No", "Shp", "Carat", "Col", "Qua", "Cut", "Pol", "Sym", "Lab", "Flo", "Sold Party", "Sale Date", "SPer", "SRate"]
    stock_headers = ["Serial No", "Shp", "Carat", "Carat_Grp", "Color name", "Qua", "Cut", "Pol", "Sym", "Lab", "Flo", "GRap", "GPer", "Expected", "Tag"]

    previous = {
        "Cart.xlsx": _xlsx_bytes(pd.DataFrame(columns=cart_headers)),
        "Hold.xlsx": _xlsx_bytes(pd.DataFrame(columns=hold_headers)),
        "Offer.xlsx": _xlsx_bytes(pd.DataFrame(columns=cart_headers)),
        "View.xlsx": _xlsx_bytes(pd.DataFrame(columns=view_headers)),
        "Wishlist.xlsx": _xlsx_bytes(pd.DataFrame(columns=cart_headers)),
        "sales_120_file.xlsx": _xlsx_bytes(pd.DataFrame(columns=sales_headers)),
        "sales_180_file.xlsx": _xlsx_bytes(pd.DataFrame(columns=sales_headers)),
        "Stock file.xlsx": _xlsx_bytes(pd.DataFrame(columns=stock_headers)),
    }
    raw = {
        "ADD_TO_CART.xlsx": _xlsx_bytes(pd.DataFrame([{"SERIAL_NO": "S1", "CREATED_DATETIME": today, "Sold Party": "P1"}])),
        "JANGAD_HISTORY.xlsx": _xlsx_bytes(pd.DataFrame([{"SERIAL_NO": "S1", "JANGAD_DATE": today, "Sold Party": "P1"}])),
        "OFFER_PLACE.xlsx": _xlsx_bytes(pd.DataFrame([{"SERIAL_NO": "S1", "CREATED_DATETIME": today, "Sold Party": "P1"}])),
        "APPOINTMENT.xlsx": _xlsx_bytes(pd.DataFrame([{"SEQ_NO": 1, "SERIAL_NO": "S1", "CREATED_DATETIME": today, "Sold Party": "P1"}])),
        "WISHLIST.xlsx": _xlsx_bytes(pd.DataFrame([{"SERIAL_NO": "S1", "CREATED_DATETIME": today, "Sold Party": "P1"}])),
        "Sales.xlsx": _xlsx_bytes(pd.DataFrame([{
            "Serial No": "S1", "Shp": "RBC", "Carat": 1.0, "Col": "G", "Qua": "VS1", "Cut": "EX",
            "Pol": "EX", "Sym": "EX", "Lab": "GIA", "Flo": "NON", "Sold Party": "P1",
            "Sale Date": today, "SPer": -10, "SRate": 100,
        }])),
        "Stock.xlsx": _xlsx_bytes(pd.DataFrame([{
            "Serial No": "S1", "Shp": "RBC", "Carat": 1.0, "Color name": "G", "Qua": "VS1", "Cut": "EX",
            "Pol": "EX", "Sym": "EX", "Lab": "GIA", "Flo": "NON", "GRap": 150, "GPer": -20, "Expected": 120,
        }])),
    }

    data, filename, logs = run_auto_pricing(
        raw_zip_bytes=_zip_bytes(raw),
        previous_final_zip_bytes=_zip_bytes(previous),
        template_bytes=_auto_pricing_template_bytes(),
    )
    assert filename == "Auto_Pricing_v4_1_3_Output.zip"
    assert "all_datasets_ok" in logs
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        names = zf.namelist()
        assert "Processing_Audit.xlsx" in names
        assert "processing_report.json" in names
        assert any(name.startswith("Final_Base_Output_") for name in names)
        final_name = next(name for name in names if name.startswith("Auto_pricing_FINAL_"))
        wb = load_workbook(io.BytesIO(zf.read(final_name)), read_only=True, data_only=True)

    ws = wb["Sheet1"]
    row = dict(zip((cell.value for cell in ws[1]), (cell.value for cell in ws[2])))
    assert row["Serial No"] == "S1"
    assert float(row["Sold_Avg_Rate_120days"]) == 100.0
    assert float(row["Diff"]) == 16.67
    assert row["Remark"] == "Yes"
    assert int(row["Unique_Buyer_Count"]) == 1
    assert int(row["sale_count"]) == 1
    assert int(row["Action_count"]) == 1
    assert int(row["ALL_Action_count"]) == 1

def test_layout_stock_matcher_adapter_uses_source_defaults_and_output_shape():
    layout = pd.DataFrame([
        {
            "Layout_ID": "L1", "Segment": "S", "SerialNo": f"L{i}",
            "Shape": "RBC", "Carat": 1.0, "Color": "G", "Clarity": "VS1",
            "Flo": "NON", "Lab": "GIA",
        }
        for i in range(1, 6)
    ])
    history = {
        "Serial No": "H1", "Shp": "RBC", "Carat": 1.0, "Col": "G",
        "Qua": "VS1", "Flo": "NON", "Lab": "GIA", "Sold Party": "Party A",
        "Country": "USA", "City": "NY", "SALE_PERSON": "Mr Test",
        "Type": "SALE", "Date": "2026-08-01",
    }
    action = {**history, "Serial No": "H2", "Type": "CART", "Date": "2026-08-02"}

    data, filename, logs = run_layout_stock_matcher(
        layout_bytes=_xlsx_bytes(layout),
        sales_bytes=_xlsx_bytes(pd.DataFrame([history])),
        actions_bytes=_xlsx_bytes(pd.DataFrame([action])),
    )
    assert filename == "Salesperson_Layout_Suggestions.zip"
    assert "'suggested_rows': 5" in logs
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        assert zf.namelist() == ["Mr Test.xlsx", "00_Run_Summary.xlsx"]
        wb = load_workbook(io.BytesIO(zf.read("Mr Test.xlsx")))
    assert wb.sheetnames == ["Navigation", "Party A"]
    assert wb["Navigation"]["C2"].value == 1
    assert wb["Party A"].max_row == 6


def test_dealmaker_package_internal_checksums_and_bundled_control():
    manifest = PACKAGE_DIR / "SHA256SUMS.txt"
    for line in manifest.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line:
            continue
        expected, filename = line.split(maxsplit=1)
        path = PACKAGE_DIR / filename.strip()
        assert path.is_file(), filename
        actual = hashlib.sha256(path.read_bytes()).hexdigest()
        assert actual == expected, filename

    assert bundled_dealmaker_control_bytes() == DEFAULT_CONTROL.read_bytes()
    assert hashlib.sha256(DEFAULT_CONTROL.read_bytes()).hexdigest() == "0797bf540f29de4df0fd94832b28e41302216bd7161fcfbfa9ab5a6534f5c288"


def test_dealmaker_production_runner_cli_is_intact():
    result = subprocess.run(
        [sys.executable, str(RUNNER), "--help"],
        cwd=str(PACKAGE_DIR),
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        check=False,
    )
    assert result.returncode == 0
    assert "--sales" in result.stdout
    assert "--actions" in result.stdout
    assert "--stock" in result.stdout
    assert "--control" in result.stdout
    assert "--output-dir" in result.stdout


def test_dealmaker_adapter_routes_default_and_override_control_without_mutating_base(monkeypatch):
    import adapters.dealmaker_adapter as adapter

    original_control = DEFAULT_CONTROL.read_bytes()
    commands: list[list[str]] = []

    class Result:
        returncode = 0
        stdout = "fake production success"

    def fake_run(command, **kwargs):
        commands.append(list(command))
        output_dir = Path(command[command.index("--output-dir") + 1])
        output_dir.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output_dir / adapter.OUTPUT_ZIP_NAME, "w") as zf:
            zf.writestr("marker.txt", "ok")
        return Result()

    monkeypatch.setattr(adapter.subprocess, "run", fake_run)

    common = dict(sales_bytes=b"sales", actions_bytes=b"actions", stock_bytes=b"stock")
    adapter.run_dealmaker(**common)
    default_command = commands[-1]
    default_control_arg = Path(default_command[default_command.index("--control") + 1])
    assert default_control_arg == DEFAULT_CONTROL

    override_bytes = b"updated-control-for-routing-test"
    adapter.run_dealmaker(**common, control_bytes=override_bytes)
    override_command = commands[-1]
    override_control_arg = Path(override_command[override_command.index("--control") + 1])
    assert override_control_arg.name == "Dealmaker_Full_Control_v6.xlsx"
    assert override_control_arg != DEFAULT_CONTROL
    assert DEFAULT_CONTROL.read_bytes() == original_control
