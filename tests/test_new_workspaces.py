from __future__ import annotations

import hashlib
import io
import subprocess
import sys
import zipfile
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from adapters.auto_pricing_adapter import ACTION_NAMES, run_auto_pricing
from adapters.dealmaker_adapter import DEFAULT_CONTROL, PACKAGE_DIR, RUNNER, bundled_dealmaker_control_bytes
from adapters.layout_stock_matcher_adapter import run_layout_stock_matcher

ROOT = Path(__file__).resolve().parents[1]


def _xlsx_bytes(frame: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    frame.to_excel(buffer, index=False)
    return buffer.getvalue()


def test_auto_pricing_preserved_source_matches_uploaded_package_copy():
    master = ROOT / "original_master_code" / "Auto_Pricing.py"
    packaged = ROOT / "original_packages" / "Auto Pricing.txt"
    assert master.read_bytes() == packaged.read_bytes()


def test_auto_pricing_adapter_preserves_core_counts_and_pricing_logic():
    tag = "RBC-1.00-1.09-G-VS1-GIA-NON"
    stock = _xlsx_bytes(pd.DataFrame([
        {"Serial No": "S1", "Tag": tag, "Carat": 1.0, "Expected": 120},
    ]))
    sales_180 = _xlsx_bytes(pd.DataFrame([
        {"Tag": tag, "Sold Party": "P1"},
        {"Tag": tag, "Sold Party": "P1"},
    ]))
    sales_120 = _xlsx_bytes(pd.DataFrame([
        {"Tag": tag, "Carat": 1.0, "Amt": 100},
    ]))
    view = _xlsx_bytes(pd.DataFrame([
        {"Serial No": "S1", "Sold Party": "P1"},
        {"Serial No": "S1", "Sold Party": "P2"},
    ]))
    action_files = {name: None for name in ACTION_NAMES}
    action_files["View"] = view

    data, filename, _ = run_auto_pricing(
        stock_bytes=stock,
        sales_180_bytes=sales_180,
        sales_120_bytes=sales_120,
        action_files=action_files,
    )
    assert filename == "Auto_Pricing_Output.zip"
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        assert zf.namelist() == [
            "Auto pricing revision.xlsx",
            "Cross_Verify_Sales_Actions.xlsx",
        ]
        output = pd.read_excel(io.BytesIO(zf.read("Auto pricing revision.xlsx")))
        verify = pd.read_excel(io.BytesIO(zf.read("Cross_Verify_Sales_Actions.xlsx")))

    row = output.iloc[0]
    assert int(row["sale_count"]) == 2
    assert int(row["Unique_Buyer_Count"]) == 1
    assert float(row["Sold_Avg_Rate_120days"]) == 100.0
    assert int(row["View_Count"]) == 1
    assert int(row["View_All_Count"]) == 2
    assert int(row["Action_count"]) == 1
    assert int(row["ALL_Action_count"]) == 2
    assert float(row["Diff"]) == 16.67
    assert row["Remark"] == "Yes"
    assert int(verify.iloc[0]["View_Count"]) == 1


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
