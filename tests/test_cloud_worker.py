from __future__ import annotations

import io
import json
import subprocess
import zipfile
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook

from adapters import suggested_adapter
from adapters.suggested_adapter import discover_salespeople_from_paths, run_suggested_batched_from_paths
from cloud_jobs.bundle import build_job_bundle
from cloud_worker.run_job import safe_extract_bundle


def _excel_bytes(headers: list[str]) -> bytes:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Sheet1"
    worksheet.append(headers)
    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def test_build_job_bundle_streams_inputs(tmp_path: Path):
    destination = tmp_path / "input_bundle.zip"
    job = build_job_bundle(
        output_path=destination,
        job_id="mmw-job-test",
        runtime_overrides={"SORT_HIGH_PRIORITY_PARTY_TYPES": {"GOLD", "VIP"}},
        scope="selected",
        salespeople=["Mr. One", "Mr. One", "Ms. Two"],
        batch_size=2,
        continue_after_failed_batch=True,
        stock_file=("stock.xlsx", io.BytesIO(_excel_bytes(["Serial No"]))),
        filter_file=("filters.xlsx", io.BytesIO(_excel_bytes(["SALES_PERSON"]))),
        requested_by="Test User",
    )
    assert job["salespeople"] == ["Mr. One", "Ms. Two"]
    with zipfile.ZipFile(destination) as archive:
        names = set(archive.namelist())
        assert {"job.json", "inputs/stock_actual.xlsx", "inputs/filters.xlsx"} <= names
        metadata = json.loads(archive.read("job.json"))
    assert metadata["scope"] == "selected"
    assert sorted(metadata["runtime_overrides"]["SORT_HIGH_PRIORITY_PARTY_TYPES"]) == ["GOLD", "VIP"]



def test_build_job_bundle_allows_repository_default_inputs(tmp_path: Path):
    destination = tmp_path / "input_bundle_defaults.zip"
    job = build_job_bundle(
        output_path=destination,
        job_id="mmw-job-defaults",
        runtime_overrides={},
        scope="selected",
        salespeople=["Mr. One"],
        batch_size=1,
        continue_after_failed_batch=True,
        stock_file=("stock.xlsx", io.BytesIO(_excel_bytes(["Serial No"]))),
        filter_file=None,
    )
    assert job["inputs"]["filters"] is None
    with zipfile.ZipFile(destination) as archive:
        assert "inputs/filters.xlsx" not in archive.namelist()


def test_salesperson_discovery_reads_profile_and_filter_columns(tmp_path: Path):
    profile_path = tmp_path / "profile.parquet"
    pd.DataFrame(
        {
            "Salesperson": ["Mr. Profile User", "Ms. Second User"],
            "Large Unused Column": ["x" * 100, "y" * 100],
        }
    ).to_parquet(profile_path, index=False)

    filter_path = tmp_path / "filters.xlsx"
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Change Format"
    worksheet.append(["SALES_PERSON", "SOLD PARTY"])
    worksheet.append(["MR. FILTER USER", "PARTY A"])
    workbook.save(filter_path)

    values = discover_salespeople_from_paths(
        profile_perm_path=profile_path,
        profile_temp_path=None,
        filter_path=filter_path,
    )
    assert values == ["Mr. Filter User", "Mr. Profile User", "Ms. Second User"]


def test_batch_wrapper_changes_only_scope_values(monkeypatch, tmp_path: Path):
    stock = tmp_path / "stock.xlsx"
    filters = tmp_path / "filters.xlsx"
    stock.write_bytes(b"stock")
    filters.write_bytes(b"filters")

    captured_runtime: list[dict[str, object]] = []

    def fake_subprocess(*, spec_path: Path, console_path: Path):
        spec = json.loads(spec_path.read_text(encoding="utf-8"))
        runtime = spec["runtime_overrides"]
        captured_runtime.append(runtime)
        output = Path(spec["output_path"])

        master_buffer = io.BytesIO()
        master = Workbook()
        master_ws = master.active
        master_ws.title = "Master_Data_Part_1"
        master_ws.append(["Salesperson", "Sold Party", "Serial No", "DETAIL", "Is_Nearby"])
        for index, person in enumerate(runtime["TEST_SALESPEOPLE"], start=1):
            master_ws.append([person, f"Party {index}", index, "DETAIL", False])
            master_ws.cell(master_ws.max_row, 4).hyperlink = f"https://example.com/{person}"
        master.save(master_buffer)

        with zipfile.ZipFile(output, "w") as archive:
            archive.writestr("MASTER_CONSOLIDATED_REPORT.xlsx", master_buffer.getvalue())
            for person in runtime["TEST_SALESPEOPLE"]:
                individual_buffer = io.BytesIO()
                individual = Workbook()
                individual.active.append(["Salesperson", person])
                individual.save(individual_buffer)
                archive.writestr(f"{person}.xlsx", individual_buffer.getvalue())
                separated_buffer = io.BytesIO()
                with zipfile.ZipFile(separated_buffer, "w") as separated:
                    separated.writestr("Party.xlsx", individual_buffer.getvalue())
                archive.writestr(f"{person}_Separated_Parties.zip", separated_buffer.getvalue())

        Path(spec["log_path"]).write_text("ok", encoding="utf-8")
        console_path.write_text("", encoding="utf-8")
        return subprocess.CompletedProcess(args=["fake"], returncode=0)

    monkeypatch.setattr(suggested_adapter, "_run_batch_subprocess", fake_subprocess)
    result = tmp_path / "result.zip"
    manifest = run_suggested_batched_from_paths(
        stock_path=stock,
        profile_perm_path=None,
        profile_temp_path=None,
        filter_path=filters,
        clean_path=None,
        sales_history_paths=[],
        runtime_overrides={
            "TOLERANCE_LEVEL": 2,
            "ENABLE_STYLING": True,
            "ENABLE_NEW_FILL_LOGIC": True,
            "GENERATE_MERGED_MASTER": True,
            "GENERATE_INDIVIDUAL_REPORTS": True,
            "ENABLE_OUTPUT_FILE_SEPARATED": True,
            # The final global master must ignore this intermediate limit.
            "MAX_ROWS_PER_SHEET": 1,
        },
        scope="selected",
        salespeople=["A", "B", "C"],
        batch_size=2,
        output_path=result,
    )

    assert manifest["final_state"] == "completed"
    assert [runtime["TEST_SALESPEOPLE"] for runtime in captured_runtime] == [["A", "B"], ["C"]]
    assert all(runtime["TEST_MODE"] is True for runtime in captured_runtime)
    assert all(runtime["TOLERANCE_LEVEL"] == 2 for runtime in captured_runtime)
    assert all(runtime["ENABLE_STYLING"] is True for runtime in captured_runtime)
    with zipfile.ZipFile(result) as archive:
        names = set(archive.namelist())
        assert "MASTER_CONSOLIDATED_REPORT.xlsx" in names
        assert {"A.xlsx", "B.xlsx", "C.xlsx"} <= names
        assert {
            "A_Separated_Parties.zip",
            "B_Separated_Parties.zip",
            "C_Separated_Parties.zip",
        } <= names
        assert "batch_01/suggested_reports.zip" not in names
        assert "_batch_diagnostics/batch_01/processing_log.txt" in names
        assert "_batch_diagnostics/batch_02/processing_log.txt" in names
        assert "batch_manifest.json" in names
        assert "processing_log.txt" in names
        manifest_payload = json.loads(archive.read("batch_manifest.json"))
        assert manifest_payload["schema_version"] == 3
        master_merge = manifest_payload["output_layout"]["master_merge"]
        assert master_merge["rows"] == 3
        assert master_merge["worksheets"] == 1
        assert master_merge["sheet_policy"] == "strict_single_sheet"
        assert master_merge["assembler_version"] == "1.4.3"
        master_bytes = archive.read("MASTER_CONSOLIDATED_REPORT.xlsx")

    merged = load_workbook(io.BytesIO(master_bytes), read_only=False, data_only=False)
    try:
        assert merged.sheetnames == ["Master_Data"]
        worksheet = merged.active
        assert worksheet.max_row == 4
        assert [worksheet.cell(row, 1).value for row in range(2, 5)] == ["A", "B", "C"]
        assert worksheet["D2"].value.startswith('=HYPERLINK("https://example.com/')
    finally:
        merged.close()


def test_safe_extract_rejects_parent_traversal(tmp_path: Path):
    source = tmp_path / "bad.zip"
    with zipfile.ZipFile(source, "w") as archive:
        archive.writestr("../escape.txt", "bad")
    try:
        safe_extract_bundle(source, tmp_path / "extract")
    except ValueError as exc:
        assert "Unsafe path" in str(exc)
    else:
        raise AssertionError("Traversal entry was not rejected")
