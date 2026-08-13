from __future__ import annotations

import hashlib
import io
import sys
import zipfile
from pathlib import Path

from openpyxl import Workbook, load_workbook

from adapters.common import MASTER_CODE_DIR, load_fresh_module
from adapters.separate_adapter import run_separate_excel_files

ROOT = Path(__file__).resolve().parents[1]


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def test_master_sources_match_repository_checksum_manifest():
    manifest = ROOT / "SOURCE_CHECKSUMS.sha256"
    lines = [line.strip() for line in manifest.read_text(encoding="utf-8").splitlines() if line.strip()]
    assert len(lines) == 12
    for line in lines:
        expected, relative = line.split(maxsplit=1)
        path = ROOT / relative.strip()
        assert path.is_file()
        assert digest(path) == expected


def test_importable_master_modules():
    modules = [
        load_fresh_module(MASTER_CODE_DIR / "Suggested_Process.py", "test_suggested"),
        load_fresh_module(MASTER_CODE_DIR / "Filter_Demand_Stock_Process.py", "test_demand"),
        load_fresh_module(MASTER_CODE_DIR / "Sorting_Suggested_Output_Zip_Process.py", "test_sorting"),
        load_fresh_module(MASTER_CODE_DIR / "BID_Drop_Checking_Process.py", "test_bid_drop"),
        load_fresh_module(MASTER_CODE_DIR / "USA_Suggested_Stock_Mapper.py", "test_usa_suggested"),
    ]
    try:
        assert callable(modules[0].generate_salesperson_reports_from_profile)
        assert callable(modules[1].generate_salesperson_reports)
        assert callable(modules[2].main)
        assert callable(modules[3].run_bid_drop_check)
        assert callable(modules[4].generate_salesperson_reports)
    finally:
        for module in modules:
            sys.modules.pop(module.__name__, None)


def test_linux_sheet_separation_adapter():
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Party A"
    ws1.append(["Serial No", "Value"])
    ws1.append(["A1", 10])
    ws1["A2"].hyperlink = "https://example.com/a1"
    ws2 = wb.create_sheet("Party B")
    ws2.append(["Serial No", "Value"])
    ws2.append(["B1", 20])

    workbook_buffer = io.BytesIO()
    wb.save(workbook_buffer)

    source_zip_buffer = io.BytesIO()
    with zipfile.ZipFile(source_zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("Salesperson.xlsx", workbook_buffer.getvalue())

    output, filename, logs = run_separate_excel_files(input_zip_bytes=source_zip_buffer.getvalue())
    assert filename.endswith(".zip")
    assert "Found 1 workbook" in logs

    with zipfile.ZipFile(io.BytesIO(output), "r") as outer:
        assert outer.namelist() == ["Salesperson.zip"]
        nested = outer.read("Salesperson.zip")
    with zipfile.ZipFile(io.BytesIO(nested), "r") as inner:
        assert set(inner.namelist()) == {"Party A.xlsx", "Party B.xlsx"}
        party_a = inner.read("Party A.xlsx")

    check = load_workbook(io.BytesIO(party_a))
    assert check.sheetnames == ["Party A"]
    assert check["Party A"]["A2"].value == "A1"
    assert check["Party A"]["A2"].hyperlink.target == "https://example.com/a1"
