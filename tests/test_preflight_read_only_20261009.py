"""Regression for Streamlit Cloud: read-only header preview must never create folders."""
import io
import os
import subprocess
import sys
from pathlib import Path
from openpyxl import Workbook
from adapters.auto_pricing_pipeline_adapter import preview_auto_pricing_column_mapping


def excel_bytes(headers):
    w = Workbook()
    w.active.append(headers)
    buf = io.BytesIO()
    w.save(buf)
    return buf.getvalue()


def test_import_processor_does_not_write_to_output_path(tmp_path):
    missing_output = tmp_path / 'not_created' / 'nested'
    root = Path(__file__).resolve().parents[1]
    env = dict(os.environ, OUTPUT_DIR=str(missing_output), PYTHONPATH=str(root))
    proc = subprocess.run([sys.executable, '-c',
        'from processors.auto_pricing_finalizer_v4_1_3.process_datasets_v4_1_3 import schema_mapping; print("OK")'],
        cwd=root, env=env, text=True, capture_output=True)
    assert proc.returncode == 0, proc.stderr
    assert 'OK' in proc.stdout
    assert not missing_output.exists()


def test_mapping_preview_read_only_no_temp_sales_file(monkeypatch, tmp_path):
    import adapters.auto_pricing_pipeline_adapter as adapter
    def forbidden_directory(*a, **kw):
        raise PermissionError('TemporaryDirectory should not be needed by header preview')
    monkeypatch.setattr(adapter.tempfile, 'TemporaryDirectory', forbidden_directory)
    data = {
        'Sales': excel_bytes(['SERIAL_NO', 'SALE_DATE', 'SRate', 'Shp', 'Carat', 'Col', 'Qua', 'Lab', 'Flo', 'Sold Party']),
        'Hold': excel_bytes(['Serial No', 'Jangad Date', 'Party Name']),
    }
    rows = preview_auto_pricing_column_mapping(data)
    date = next(r for r in rows if r['Dataset'] == 'Sales' and r['Final column'] == 'Sale Date')
    assert date['Source column'] == 'SALE_DATE' and date['Status'] == 'OK'
    party_code = next(r for r in rows if r['Dataset'] == 'Hold' and r['Final column'] == 'PARTY_CODE')
    assert party_code['Status'] == 'OPTIONAL_MISSING'
    assert not (tmp_path/'sales_map_preview_').exists()


def test_processor_import_forbidden_output_location():
    root = Path(__file__).resolve().parents[1]
    env = dict(os.environ, OUTPUT_DIR='/proc/auto_pricing_readonly_should_not_exist', PYTHONPATH=str(root))
    proc = subprocess.run([sys.executable, '-c',
        'from processors.auto_pricing_finalizer_v4_1_3.process_datasets_v4_1_3 import schema_mapping; print("OK")'],
        cwd=root, env=env, text=True, capture_output=True)
    assert proc.returncode == 0, proc.stderr
    assert 'OK' in proc.stdout
