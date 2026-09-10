from pathlib import Path


def test_app_contains_all_workspaces():
    text = (Path(__file__).resolve().parents[1] / "app.py").read_text(encoding="utf-8")
    for title in [
        "Suggested Stock Mapper",
        "USA Suggested Stock Mapper",
        "Auto Pricing",
        "Dealmaker",
        "Layout Stock Matcher",
        "Demand Filter Mapper",
        "Sort Suggested Output",
        "Suggested List Splitter",
        "BID Drop Checking Process",
        "Separate Excel Sheets",
    ]:
        assert title in text
    assert "Developed by Keyur Bhojak" in (Path(__file__).resolve().parents[1] / "ui" / "theme.py").read_text()


def test_suggested_page_exposes_full_master_controls():
    text = (Path(__file__).resolve().parents[1] / "app.py").read_text(encoding="utf-8")
    required_controls = [
        "SHOW_BASIC_SUPERTAG",
        "SHOW_ADVANCED_SUPERTAG",
        "SHOW_REF_STONE_ID",
        "SHOW_CONCAT_STATUS",
        "GENERATE_INDIVIDUAL_REPORTS",
        "GENERATE_MERGED_MASTER",
        "ENABLE_TYPE_COLUMNS",
        "MATCH_STATUS_FILTER_MODE",
        "ENABLE_TOP_CUSTOMERS_HIGHLIGHT",
        "SORT_COUNT_THRESHOLD",
        "TEST_SALESPEOPLE",
        "TEST_PARTIES",
    ]
    for control in required_controls:
        assert control in text



def test_all_suggested_user_input_constants_are_wired():
    text = (Path(__file__).resolve().parents[1] / "app.py").read_text(encoding="utf-8")
    direct_expected = {
        "TOP_N_SUPERTAGS", "TOLERANCE_LEVEL", "ENABLE_STYLING", "ENABLE_FLO_MAPPING",
        "SHOW_BASIC_SUPERTAG", "SHOW_ADVANCED_SUPERTAG", "SHOW_REF_STONE_ID",
        "SHOW_CONCAT_STATUS", "GENERATE_INDIVIDUAL_REPORTS", "GENERATE_MERGED_MASTER",
        "ENABLE_CLEAN_FILE", "ENABLE_PAIR_MODE", "ENABLE_TYPE_COLUMNS",
        "STATIC_TYPE_SEQUENCE", "ENABLE_SALES_SUPERTAG_COLUMN", "MAX_ROWS_PER_SHEET",
        "TEST_MODE", "TEST_SALESPEOPLE", "TEST_MODE_PARTY", "TEST_PARTIES",
        "MATCH_STATUS_FILTER_MODE", "MATCH_STATUS_ALLOWED_TYPES",
        "MATCH_STATUS_SPECIFIC_PARTIES", "ENABLE_NEW_FILL_LOGIC",
        "ENABLE_TOP_CUSTOMERS_HIGHLIGHT", "TOP_CUSTOMERS_N",
        "ENABLE_SMALL_OUTPUT_DELETE", "SMALL_OUTPUT_THRESHOLD",
        "ENABLE_OUTPUT_FILE_SEPARATED", "ENABLE_ADVANCED_SORTING",
        "SORT_COUNT_THRESHOLD", "SORT_NORMAL_PARTY_LIMIT", "SORT_HIGH_PRIORITY_PARTY_LIMIT",
        "SORT_EXACT_CAP", "SORT_NEAR_CAP", "SORT_MAYBE_CAP",
        "SORT_HIGH_PRIORITY_PARTY_TYPES", "SORT_HIGH_PRIORITY_TOTAL_HISTORY_THRESHOLD",
    }
    missing = sorted(name for name in direct_expected if name not in text)
    assert not missing, missing
    assert 'runtime[f"INCLUDE_{label}"]' in text
    assert 'runtime[f"ENABLE_{label}_TOLERANCE"]' in text
    for label in [
        "COLOR", "CLARITY", "FLORO", "CUT", "POLISH", "SYMMETRY", "SHADE",
        "MILKY", "TABLE_BLACK", "SIDE_BLACK", "TABLE_OPEN", "SIDE_OPEN",
        "GIRDLE_OPEN", "PAVILION_OPEN",
    ]:
        assert f'"{label}":' in text


def test_completed_cloud_job_keeps_final_batch_progress():
    root = Path(__file__).resolve().parents[1]
    app_text = (root / "app.py").read_text(encoding="utf-8")
    worker_text = (root / "cloud_worker" / "run_job.py").read_text(encoding="utf-8")
    assert 'current_batch = total_batches' in app_text
    assert '"current_batch": total_batches' in worker_text
    assert '"stage": "completed"' in worker_text


def test_cloud_result_describes_combined_output_layout():
    text = (Path(__file__).resolve().parents[1] / "app.py").read_text(encoding="utf-8")
    assert "one global MASTER_CONSOLIDATED_REPORT.xlsx" in text
    assert "Download final combined result ZIP" in text


def test_deployment_exposes_version_markers():
    root = Path(__file__).resolve().parents[1]
    app_text = (root / "app.py").read_text(encoding="utf-8")
    worker_text = (root / "cloud_worker" / "run_job.py").read_text(encoding="utf-8")
    workflow_text = (root / ".github" / "workflows" / "process-suggested-cloud.yml").read_text(encoding="utf-8")
    assert "APP_VERSION" in app_text
    assert "Application version" not in app_text
    assert '"worker_version": WORKER_VERSION' in worker_text
    assert "Verify deployed project version" in workflow_text


def test_auto_pricing_full_pipeline_ui_is_exposed():
    root = Path(__file__).resolve().parents[1]
    text = (root / "app.py").read_text(encoding="utf-8")
    section = text[text.index("def auto_pricing_page()"):text.index("def dealmaker_page()") ]
    assert "Create Latest Output" in section
    assert "Latest files" in section
    assert "Previous files" in section
    assert "Stock file *" in section
    assert "Sales file *" in section
    assert "Previous 180-day Sales file *" in section
    assert "View file (optional)" in section
    assert "Previous action files (optional)" in section
    assert "Removed parties list" in section
    assert "Files ready for this run" in section
    assert "Files used and results" in section
    assert "Auto Pricing template" not in section
    assert "Git" not in section
    assert "RAW" not in section
    assert "v4.1.3" not in section
    assert "Previous Stock" not in section
    assert "Create Auto Pricing" in section


def test_auto_pricing_saved_files_exist():
    root = Path(__file__).resolve().parents[1]
    assert (root / "data" / "remove_parties.xlsx").is_file()
    assert (root / "processors" / "auto_pricing_finalizer_v4_1_3" / "process_datasets_v4_1_3.py").is_file()
