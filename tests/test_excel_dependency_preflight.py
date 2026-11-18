"""Guard the two demand-filter workspaces against missing Excel packages."""
from pathlib import Path
from unittest.mock import patch
import sys

from ui.excel_dependencies import excel_dependency_error


ROOT = Path(__file__).resolve().parents[1]


def test_openpyxl_can_be_imported_for_demand_filters():
    assert excel_dependency_error() is None


def test_missing_openpyxl_has_clear_error():
    with patch.dict(sys.modules, {"openpyxl": None}):
        detail = excel_dependency_error()
    assert detail and "openpyxl" in detail


def test_both_demand_filter_workspaces_check_excel_runtime():
    app = (ROOT / "app.py").read_text(encoding="utf-8")
    standard = app.split("def demand_filter_page() -> None:", 1)[1].split("def advanced_demand_filter_page() -> None:", 1)[0]
    advanced = app.split("def advanced_demand_filter_page() -> None:", 1)[1].split("def sorting_page() -> None:", 1)[0]
    assert "if not _excel_processing_ready():" in standard
    assert "if not _excel_processing_ready():" in advanced
    assert "openpyxl==3.1.5" in (ROOT / "requirements.txt").read_text(encoding="utf-8")
