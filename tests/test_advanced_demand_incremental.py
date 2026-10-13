"""Checkpoint runner: preserve output and resume across independent invocations."""
import io
import zipfile

import pandas as pd
import pytest
from openpyxl import load_workbook

from adapters import advanced_demand_checkpoint_adapter as controller
from adapters.advanced_demand_filter_adapter import run_advanced_demand_filter
from tests.test_advanced_demand_filter import _multi_party_stock_and_filters, _stock_bytes


def _excel_entries(payload):
    result = {}
    with zipfile.ZipFile(io.BytesIO(payload)) as z:
        for name in z.namelist():
            if name.endswith('.zip'):
                with zipfile.ZipFile(io.BytesIO(z.read(name))) as nested:
                    for nested_name in nested.namelist():
                        result[(name, nested_name)] = nested.read(nested_name)
            else:
                result[(name,)] = z.read(name)
    return result


@pytest.mark.parametrize('layout', ['salesperson_workbooks', 'separate_party_files', 'both'])
def test_checkpointed_matches_legacy_for_every_layout_and_restart(tmp_path, monkeypatch, layout):
    monkeypatch.setattr(controller, 'JOBS_ROOT', tmp_path)
    stock, filters = _multi_party_stock_and_filters()
    baseline, _, _ = run_advanced_demand_filter(
        stock_bytes=stock, filter_bytes=filters, min_stones_per_party=2,
        output_layout=layout, runtime_overrides={'TOLERANCE_LEVEL': 0},
    )
    job_id = controller.create_batch_job(
        stock_bytes=stock, filter_bytes=filters, min_stones_per_party=2,
        output_layout=layout, batch_rows=2, runtime_overrides={'TOLERANCE_LEVEL': 0},
    )
    assert controller.batch_job_snapshot(job_id)[1]['phase'] == 'ready'
    states = []
    for _ in range(30):
        # Every step reloads the code and state from the disk (independent of
        # previous calls), simulating Streamlit reruns/reconnects.
        state = controller.process_next_batch(job_id)
        states.append(state['phase'])
        if state['phase'] == 'done':
            break
    assert 'matching' in states and 'reporting' in states and 'packaging' in states
    assert states[-1] == 'done'
    assert state['filter_offset'] == 7
    result = controller.batch_output_path(job_id).read_bytes()
    original_files, batched_files = _excel_entries(baseline), _excel_entries(result)
    assert original_files.keys() == batched_files.keys()
    for path in original_files:
        a, b = load_workbook(io.BytesIO(original_files[path])), load_workbook(io.BytesIO(batched_files[path]))
        try:
            assert a.sheetnames == b.sheetnames
            for sheet in a.sheetnames:
                assert list(a[sheet].values) == list(b[sheet].values)
                for row in a[sheet].iter_rows():
                    for cell in row:
                        second = b[sheet][cell.coordinate]
                        assert (cell.hyperlink.target if cell.hyperlink else None) == (
                            second.hyperlink.target if second.hyperlink else None
                        )
                        for prop in ('font','fill','border','alignment','number_format'):
                            assert str(getattr(cell, prop)) == str(getattr(second, prop))
        finally:
            a.close()
            b.close()


def test_checkpoint_dedup_crosses_batches_and_respects_status_and_removed(tmp_path, monkeypatch):
    monkeypatch.setattr(controller, 'JOBS_ROOT', tmp_path)
    filters = pd.DataFrame([
        {'Sold Party':'LEFT', 'Sales_Person':'Agent','Shp':'RBC','Remarks':'INCLUDE'},
        {'Sold Party':'LEFT', 'Sales_Person':'Agent','Shp':'RBC','Remarks':'INCLUDE'},
        {'Sold Party':'REMOVED', 'Sales_Person':'Agent','Shp':'RBC','Remarks':'INCLUDE'},
        {'Sold Party':'RIGHT', 'Sales_Person':'Agent','Shp':'RBC','Remarks':'INCLUDE'},
    ])
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as w:
        pd.DataFrame({'Sold Party':['REMOVED']}).to_excel(w, sheet_name='Removed Party',index=False)
        filters.to_excel(w,sheet_name='Change Format',index=False)
    overrides={'TOLERANCE_LEVEL':0,'MATCH_STATUS_FILTER_MODE':'ALL',
               'MATCH_STATUS_ALLOWED_TYPES':['Exact','Near by']}
    baseline, _, _ = run_advanced_demand_filter(
        stock_bytes=_stock_bytes(),filter_bytes=out.getvalue(),runtime_overrides=overrides
    )
    job_id = controller.create_batch_job(stock_bytes=_stock_bytes(),filter_bytes=out.getvalue(),
                                          runtime_overrides=overrides,batch_rows=1)
    for _ in range(15):
        state = controller.process_next_batch(job_id)
        if state['phase'] == 'done':
            break
    assert state['phase'] == 'done'
    assert state['removed_filter_rows'] == 1
    assert state['duplicate_filter_rows'] == 1
    assert state['processed_filter_rows'] == 2
    files_a = _excel_entries(baseline)
    files_b = _excel_entries(controller.batch_output_path(job_id).read_bytes())
    assert files_a.keys() == files_b.keys()
    for path in files_a:
        a=load_workbook(io.BytesIO(files_a[path])); b=load_workbook(io.BytesIO(files_b[path]))
        assert a.sheetnames == b.sheetnames
        for tab in a.sheetnames:
            assert list(a[tab].values) == list(b[tab].values)
        a.close(); b.close()


def test_invalid_resume_id_and_missing_job(tmp_path,monkeypatch):
    monkeypatch.setattr(controller,'JOBS_ROOT', tmp_path)
    with pytest.raises(ValueError):
        controller.batch_job_snapshot('../../evil')
    with pytest.raises(FileNotFoundError):
        controller.batch_job_snapshot('0'*40)


def test_checkpointed_clean_exclusion_matches_legacy(tmp_path, monkeypatch):
    monkeypatch.setattr(controller, 'JOBS_ROOT', tmp_path)
    stock, filters = _multi_party_stock_and_filters()
    no_need = io.BytesIO()
    with pd.ExcelWriter(no_need, engine='openpyxl') as w:
        pd.DataFrame([{
            'Sales_Person': 'Mr. First Seller',
            'Sold Party': 'TWO STONES',
            'Remarks': 'INCLUDE',
            'COLOR NAME': 'D',
        }]).to_excel(w, sheet_name='Sheet1', index=False)
    clean = no_need.getvalue()
    baseline, _, _ = run_advanced_demand_filter(
        stock_bytes=stock, filter_bytes=filters, clean_bytes=clean,
        runtime_overrides={'TOLERANCE_LEVEL': 0},
    )
    job_id = controller.create_batch_job(
        stock_bytes=stock, filter_bytes=filters, clean_bytes=clean,
        batch_rows=3, runtime_overrides={'TOLERANCE_LEVEL': 0},
    )
    for _ in range(20):
        state = controller.process_next_batch(job_id)
        if state['phase'] == 'done':
            break
    assert state['phase'] == 'done'
    original_files = _excel_entries(baseline)
    batch_files = _excel_entries(controller.batch_output_path(job_id).read_bytes())
    assert original_files.keys() == batch_files.keys()
    for path in original_files:
        a=load_workbook(io.BytesIO(original_files[path])); b=load_workbook(io.BytesIO(batch_files[path]))
        for tab in a.sheetnames:
            assert list(a[tab].values)==list(b[tab].values)
        a.close(); b.close()


def test_failing_batch_does_not_skip_rows_or_keep_partial_matches(tmp_path, monkeypatch):
    """An exception before checkpoint must roll back SQLite and cursor together."""
    from adapters.common import load_fresh_module, MASTER_CODE_DIR
    from processors.advanced_demand_incremental import run_checkpoint_step, get_state
    monkeypatch.setattr(controller, 'JOBS_ROOT', tmp_path)
    stock, filters = _multi_party_stock_and_filters()
    job_id = controller.create_batch_job(stock_bytes=stock, filter_bytes=filters, batch_rows=3)
    directory = controller._job_folder(job_id)
    master = load_fresh_module(MASTER_CODE_DIR/'Advance_Demand_Filter_Process.py', 'checkpoint_fault_master')
    real_match = master.apply_tolerance_filter
    calls = {'n':0}
    def fail_mid_batch(*args, **kwargs):
        calls['n'] += 1
        if calls['n'] == 2:
            raise RuntimeError('simulated failure in uncommitted batch')
        return real_match(*args, **kwargs)
    master.apply_tolerance_filter = fail_mid_batch
    with pytest.raises(RuntimeError, match='simulated failure'):
        run_checkpoint_step(master, str(directory/'stock.xlsx'),
                            str(directory/'failed.zip'), str(directory/'filters.xlsx'),
                            1, job_dir=directory, batch_rows=3)
    state = get_state(directory)
    assert state['filter_offset'] == 0
    assert state['processed_filter_rows'] == 0
    for _ in range(25):
        state = controller.process_next_batch(job_id)
        if state['phase'] == 'done':
            break
    assert state['phase'] == 'done'
    assert state['processed_filter_rows'] == 7


def test_time_limited_batches_preserve_cursor_and_complete(tmp_path):
    from adapters.common import load_fresh_module, MASTER_CODE_DIR
    from processors.advanced_demand_incremental import run_checkpoint_step
    stock, filters = _multi_party_stock_and_filters()
    (tmp_path/'stock.xlsx').write_bytes(stock)
    (tmp_path/'filters.xlsx').write_bytes(filters)
    for i in range(30):
        master = load_fresh_module(MASTER_CODE_DIR/'Advance_Demand_Filter_Process.py',f'checkpoint_time_{i}')
        state = run_checkpoint_step(master, str(tmp_path/'stock.xlsx'),
                                    str(tmp_path/'output.zip'), str(tmp_path/'filters.xlsx'),
                                    1, job_dir=tmp_path, batch_rows=100,
                                    max_batch_seconds=0)
        if state['phase'] == 'done':
            break
    assert state['phase'] == 'done'
    assert state['filter_offset'] == 7
    assert state['processed_filter_rows'] == 7
