"""October 2026 vendor export compatibility and no-silent-data-loss mapping checks."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook

from processors.auto_pricing_finalizer_v4_1_3 import process_datasets_v4_1_3 as p
from adapters.auto_pricing_pipeline_adapter import ACTION_HEADERS, preview_auto_pricing_column_mapping


def _book(tmp_path, name, headers, rows):
    wb = Workbook()
    ws = wb.active
    ws.append(headers)
    for row in rows:
        ws.append(row)
    path = tmp_path / name
    wb.save(path)
    return path


def _source(row, dataset, headers):
    rec, _ = p.map_row(row, dataset, headers)
    return rec


def test_current_october_offer_wishlist_date_and_business_fields():
    raw = {
        'Created Date': datetime(2026, 10, 8, 12, 19),
        'Color': 'D', 'Quality': 'VVS1',
        'TBInc': 'T1', 'TOIn': 'T2', 'SBInc': 'S1', 'SOIn': 'S2',
        'Girdle Open': 'G1', 'Pav Open': 'P1',
    }
    for dataset in ('Offer', 'Wishlist', 'Cart'):
        mapped = _source(raw, dataset, ACTION_HEADERS[dataset])
        assert mapped['CREATED_DATETIME'] == raw['Created Date']
        for final, expected in [('Col','D'), ('Qua','VVS1'), ('TBINC_NAME','T1'),
                                ('TOIN_NAME','T2'),('SBINC_NAME','S1'),('SOIN_NAME','S2'),
                                ('GIRDLE_OPEN_NAME','G1'),('PAV_OPEN_NAME','P1')]:
            assert mapped[final] == expected
        mapping = p.schema_mapping(dataset, ACTION_HEADERS[dataset], list(raw))
        relevant = [r for r in mapping if r['Final Column'] in ('CREATED_DATETIME','Col','Qua')]
        assert all(r['Status'] == 'OK' for r in relevant)


def test_old_export_mapping_is_unchanged():
    rec = {'CREATED_DATETIME': datetime(2026, 10, 2, 1, 0), 'Col':'G', 'Qua':'VS1',
           'TBINC_NAME':'A', 'GIRDLE_OPEN_NAME':'B'}
    out = _source(rec, 'Offer', ACTION_HEADERS['Offer'])
    for target, value in rec.items():
        assert out[target] == value


def test_hold_party_code_is_not_guessed_from_party_name():
    columns = ['Serial No', 'Jangad Date', 'Party Name']
    mapped = p.schema_mapping('Hold', ACTION_HEADERS['Hold'], columns)
    status = {r['Final Column']:r['Status'] for r in mapped}
    assert status['JANGAD_DATE'] == 'OK'
    assert status['Sold Party'] == 'OK'
    assert status['SERIAL_NO'] == 'OK'
    assert status['PARTY_CODE'] == 'OPTIONAL_MISSING'
    out = _source({'Party Name':'SOME PARTY', 'Jangad Date':datetime(2026,10,8)},
                  'Hold', ACTION_HEADERS['Hold'])
    assert out['PARTY_CODE'] is None


def test_stock_expected_can_be_derived_only_with_valid_operands(tmp_path, monkeypatch):
    monkeypatch.setattr(p, 'OUT_DIR', tmp_path)
    raw = _book(tmp_path, 'KSTAKA.xlsx',
                ['Serial No','Shp','Carat','Col','Qua','Cut','Pol','Sym','Lab','Flo','GRap','GPer'],
                [[100,'RBC',1.0,'D','VVS1','EX','EX','EX','GIA','NON',2000,20],
                 [101,'RBC',1.0,'D','VVS1','EX','EX','EX','GIA','NON',4000,30]])
    audit = {'summary':[],'mapping':[],'blanks':[],'issues':[]}
    result, ok = p.process_stock(raw, audit=audit)
    assert ok
    assert result['Status'] == 'PASS'
    assert next(r for r in audit['mapping'] if r['Final Column']=='Expected')['Status']=='DERIVED'
    out = load_workbook(tmp_path/'Stock file.xlsx', read_only=True, data_only=True)
    expected = [r[13] for r in list(out.active.values)[1:]]
    assert sorted(round(x,2) for x in expected) == [1350, 2340]
    out.close()
    broken = _book(tmp_path,'bad_stock.xlsx',
                   ['Serial No','Shp','Carat','Col','Qua','Cut','Pol','Sym','Lab','Flo','GRap','GPer'],
                   [[900,'RBC',1,'G','SI1','EX','EX','EX','GIA','NON',2000,'NOT-A-NUMBER']])
    audit = {'summary':[],'mapping':[],'blanks':[],'issues':[]}
    result, ok = p.process_stock(broken, 'bad_output.xlsx', audit=audit)
    assert not ok
    assert result['Stock Expected Invalid Rows'] == 1
    assert any(r['Issue Type']=='Expected cannot be calculated safely' for r in audit['issues'])


def test_offer_date_mapping_run_keeps_rows(tmp_path, monkeypatch):
    monkeypatch.setattr(p,'OUT_DIR',tmp_path)
    monkeypatch.setenv('TEST_PROCESSING_DATE','2026-10-09')
    headers = ACTION_HEADERS['Offer']
    history = _book(tmp_path, 'prev_offer.xlsx', headers, [])
    source = _book(tmp_path,'raw_offer.xlsx',
                   ['Created Date','Sold Party','Serial No','Color','Quality','TBInc','TOIn','SBInc','SOIn','Girdle Open','Pav Open'],
                   [[datetime(2026,10,8,8,5),'A',111,'E','VS1','T','O','B','S','G','P'],
                    [datetime(2026,10,8,9,5),'B',112,'D','VVS2','T','O','B','S','G','P']])
    audit = {'summary':[],'mapping':[],'blanks':[],'issues':[]}
    # Other business fields are deliberately absent; date must still be parsed and records preserved.
    result, ok = p.process_history_log('Offer',source,history,'out_offer.xlsx','CREATED_DATETIME',90,audit=audit)
    assert result['Raw Valid Dates'] == 2
    assert result['Raw Invalid Dates'] == 0
    assert result['Output Rows'] == 2
    assert ok  # descriptive unmapped source fields are non-blocking
    assert any(i['Severity']=='WARNING' and i['Issue Type']=='Optional source column missing' for i in audit['issues'])
    out = load_workbook(tmp_path/'out_offer.xlsx',read_only=True,data_only=True)
    assert [r[0] for r in list(out.active.values)[1:]] == [datetime(2026,10,8,8,5),datetime(2026,10,8,9,5)]
    out.close()


def test_invalid_previous_view_date_is_not_silently_discarded(tmp_path, monkeypatch):
    monkeypatch.setattr(p,'OUT_DIR',tmp_path)
    monkeypatch.setenv('TEST_PROCESSING_DATE','2026-10-09')
    headers = ACTION_HEADERS['View']
    old = _book(tmp_path,'prev_view.xlsx',headers,[['x' if h=='APPOINTMENT_DATE' else 5 if h=='SERIAL_NO' else None for h in headers]])
    new = _book(tmp_path,'new_view.xlsx',headers,[[datetime(2026,10,8) if h=='APPOINTMENT_DATE' else 5 if h=='SERIAL_NO' else None for h in headers]])
    audit = {'summary':[],'mapping':[],'blanks':[],'issues':[]}
    report, ok = p.process_view(new,old,audit=audit)
    assert not ok
    assert '1 invalid history date' in report['History Invalid Dates']


def test_preview_shows_business_date_and_missing_without_substitution(tmp_path):
    source = _book(tmp_path,'preview_offer.xlsx',['Created Date','Color','Quality'],[[datetime(2026,10,9),'D','VS1']])
    hold = _book(tmp_path,'preview_hold.xlsx',['Serial No','Jangad Date','Party Name'],[[200,datetime(2026,10,9),'A']])
    rows = preview_auto_pricing_column_mapping({'Offer':source.read_bytes(),'Hold':hold.read_bytes()})
    assert any(r['Dataset']=='Offer' and r['Final column']=='CREATED_DATETIME'
               and r['Source column']=='Created Date' and r['Role']=='Business date (window)' for r in rows)
    assert any(r['Dataset']=='Hold' and r['Final column']=='PARTY_CODE' and r['Status']=='OPTIONAL_MISSING' for r in rows)
