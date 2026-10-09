"""Only pricing-impacting fields block the finalizer; optional columns retain blank output."""
from datetime import datetime
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from adapters.auto_pricing_pipeline_adapter import ACTION_HEADERS, preview_auto_pricing_column_mapping
from processors.auto_pricing_finalizer_v4_1_3 import process_datasets_v4_1_3 as p


def book(tmp_path: Path, filename: str, headings, rows):
    wb = Workbook()
    ws = wb.active
    ws.append(list(headings))
    for row in rows:
        ws.append(list(row))
    out = tmp_path / filename
    wb.save(out)
    return out


def audit():
    return {'summary': [], 'mapping': [], 'blanks': [], 'issues': []}


def status(dataset, fields, all_headers):
    return {r['Final Column']: r['Status'] for r in p.schema_mapping(dataset, all_headers, fields)}


@pytest.mark.parametrize('dataset,date_field,serial_field', [
    ('Cart','CREATED_DATETIME','Serial No'),('Offer','CREATED_DATETIME','Serial No'),
    ('Wishlist','CREATED_DATETIME','Serial No'),('Hold','JANGAD_DATE','SERIAL_NO'),
    ('View','APPOINTMENT_DATE','SERIAL_NO'),
])
def test_action_pricing_minimal_fields_pass_header_preflight(dataset,date_field,serial_field):
    fields=[date_field,serial_field,'Sold Party']
    statuses=status(dataset,fields,ACTION_HEADERS[dataset])
    assert not [k for k,v in statuses.items() if v == 'MISSING']
    assert statuses[date_field]=='OK'
    if dataset == 'Hold':
        assert statuses['PARTY_CODE']=='OPTIONAL_MISSING'
    else:
        assert len([v for v in statuses.values() if v == 'OPTIONAL_MISSING']) >= 1


@pytest.mark.parametrize('dataset,date_field,serial_field', [
    ('Cart','CREATED_DATETIME','Serial No'),('Offer','CREATED_DATETIME','Serial No'),
    ('Wishlist','CREATED_DATETIME','Serial No'),('Hold','JANGAD_DATE','SERIAL_NO'),
    ('View','APPOINTMENT_DATE','SERIAL_NO'),
])
def test_missing_business_date_remains_required(dataset,date_field,serial_field):
    statuses=status(dataset,[serial_field,'Sold Party'],ACTION_HEADERS[dataset])
    assert statuses[date_field] == 'MISSING'


def test_hold_process_does_not_block_on_missing_party_code(tmp_path, monkeypatch):
    monkeypatch.setattr(p, 'OUT_DIR', tmp_path)
    monkeypatch.setenv('TEST_PROCESSING_DATE','2026-10-09')
    prev=book(tmp_path,'Hold_prev.xlsx',ACTION_HEADERS['Hold'],[])
    src=book(tmp_path,'JANGAD_HISTORY.xlsx',['Serial No','Jangad Date','Party Name'],[
        [101,datetime(2026,10,8,10,2),'Unknown New Party']])
    a=audit()
    summary,ok=p.process_history_log('Hold',src,prev,'Hold_out.xlsx','JANGAD_DATE',90,True,a)
    assert ok and summary['Status']=='PASS'
    assert summary['Output Rows']==1
    assert not [i for i in a['issues'] if i['Severity']=='ERROR']
    assert any(i['Issue Type']=='Optional Hold PARTY_CODE not verified' for i in a['issues'])
    wb=load_workbook(tmp_path/'Hold_out.xlsx',read_only=True,data_only=True)
    headers=next(wb.active.values)
    values=next(iter(list(wb.active.values)[1:]))
    row=dict(zip(headers,values))
    assert row['PARTY_CODE'] is None and row['Sold Party']=='Unknown New Party'
    wb.close()


def test_stock_expected_supplied_allows_missing_rap_discount_and_style_columns(tmp_path, monkeypatch):
    monkeypatch.setattr(p,'OUT_DIR',tmp_path)
    src=book(tmp_path,'stock_minimal.xlsx',
             ['Serial No','Shp','Carat','Color','Qua','Lab','Flo','Expected'],
             [[42,'RD',1.52,'F','VS1','GIA','NON',7800]])
    a=audit()
    result,ok=p.process_stock(src,'Stock_test.xlsx',a)
    assert ok and result['Stock Expected Invalid Rows']==0
    assert any(m['Final Column']=='GRap' and m['Status']=='OPTIONAL_MISSING' for m in a['mapping'])
    assert any(m['Final Column']=='Cut' and m['Status']=='OPTIONAL_MISSING' for m in a['mapping'])
    wb=load_workbook(tmp_path/'Stock_test.xlsx',read_only=True,data_only=True)
    vals=list(wb.active.values)
    row=dict(zip(vals[0],vals[1])); wb.close()
    assert row['Expected']==7800 and row['GRap'] is None and row['Cut'] is None
    assert row['Tag']=='RD-1.50-1.69-F-VS1-GIA-NON'


def test_stock_missing_expected_without_operands_stays_fatal():
    statuses=status('Stock',['Serial No','Shp','Carat','Color','Qua','Lab','Flo'],p.STOCK_FINAL_HEADERS)
    assert statuses['Expected']=='MISSING'
    assert statuses['GRap']=='MISSING' and statuses['GPer']=='MISSING'


def test_sales_unused_style_columns_do_not_block(tmp_path,monkeypatch):
    monkeypatch.setattr(p,'OUT_DIR',tmp_path)
    monkeypatch.setenv('TEST_PROCESSING_DATE','2026-10-09')
    critical=['Serial No','Shp','Carat','Col','Qua','Lab','Flo','Sold Party','Sale Date','SRate']
    prev=book(tmp_path,'Sales_prev_120.xlsx',p.STANDARD_SALES_HEADERS,[])
    prev180=book(tmp_path,'Sales_prev_180.xlsx',p.STANDARD_SALES_HEADERS,[])
    source=book(tmp_path,'SALES_RAW.xlsx',critical,[[101,'RD',1.0,'F','VS1','GIA','NON','Buyer A',datetime(2026,10,8),9200]])
    a=audit()
    reports,ok=p.process_sales(source,prev,prev180,a)
    assert ok and all(x['Status']=='PASS' for x in reports)
    assert not [i for i in a['issues'] if i['Severity']=='ERROR']
    assert any(m['Final Column']=='SPer' and m['Status']=='OPTIONAL_MISSING' for m in a['mapping'])
    wb=load_workbook(tmp_path/'sales_120_file.xlsx',read_only=True,data_only=True)
    vals=list(wb.active.values); row=dict(zip(vals[0],vals[1])); wb.close()
    assert row['SRate']==9200 and row['Amt']==9200
    assert row['Cut'] is None and row['SPer'] is None


def test_sales_missing_rate_still_blocks(tmp_path,monkeypatch):
    monkeypatch.setattr(p,'OUT_DIR',tmp_path)
    source=book(tmp_path,'sales_bad.xlsx',['Serial No','Shp','Carat','Col','Qua','Lab','Flo','Sold Party','Sale Date'],
                [[101,'RD',1,'F','VS1','GIA','NON','Buyer A',datetime(2026,10,8)]])
    prev=book(tmp_path,'prev1.xlsx',p.STANDARD_SALES_HEADERS,[])
    prev180=book(tmp_path,'prev2.xlsx',p.STANDARD_SALES_HEADERS,[])
    a=audit()
    reports,ok=p.process_sales(source,prev,prev180,a)
    assert not ok and all(r['Status']=='FAILED' for r in reports)
    assert any(m['Final Column']=='SRate' and m['Status']=='MISSING' for m in a['mapping'])


def test_preview_pricing_impact_marks_optional_hold_code(tmp_path):
    source=book(tmp_path,'hold.xlsx',['Serial No','Jangad Date','Party Name'],
                [[100,datetime(2026,10,8),'Buyer A']])
    rows=preview_auto_pricing_column_mapping({'Hold':source.read_bytes()})
    code=next(r for r in rows if r['Final column']=='PARTY_CODE')
    assert code['Status']=='OPTIONAL_MISSING' and code['Pricing impact']=='OPTIONAL'
    date=next(r for r in rows if r['Final column']=='JANGAD_DATE')
    assert date['Status']=='OK' and date['Pricing impact']=='REQUIRED'
