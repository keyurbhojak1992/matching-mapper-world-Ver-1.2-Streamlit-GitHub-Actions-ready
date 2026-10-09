"""Dynamic header standardization: deterministic approved aliases plus safe manual mapping."""
import io
from datetime import datetime
import json
import pytest
from openpyxl import Workbook
from processors.auto_pricing_finalizer_v4_1_3 import process_datasets_v4_1_3 as p
from adapters.auto_pricing_pipeline_adapter import preview_auto_pricing_column_mapping

def test_real_stock_exp_amt_maps_to_expected():
    m, rows = p.resolve_standard_mapping('Stock',p.STOCK_FINAL_HEADERS,['Serial No','Shp','Carat','Color name','Qua','Cut','Pol','Sym','Lab','Flo','GRap','GPer','EXP AMT'])
    assert m['Expected']=='EXP AMT'
    assert next(r for r in rows if r['Final Column']=='Expected')['Status']=='OK'

def test_unknown_header_is_not_autoselected_and_can_be_approved():
    source={'EmailDateProvided':datetime(2026,10,9), 'Stone Id':100}
    m,rows=p.resolve_standard_mapping('Offer',['CREATED_DATETIME','Serial No'],list(source))
    assert 'CREATED_DATETIME' not in m
    assert next(r for r in rows if r['Final Column']=='CREATED_DATETIME')['Status']=='MISSING'
    m,rows=p.resolve_standard_mapping('Offer',['CREATED_DATETIME','Serial No'],list(source), {'Offer':{'CREATED_DATETIME':'EmailDateProvided'}})
    assert m['CREATED_DATETIME']=='EmailDateProvided'
    mapped,_=p.map_row(source,'Offer',['CREATED_DATETIME','Serial No'],m)
    assert mapped['CREATED_DATETIME']==source['EmailDateProvided']

def test_no_wrong_party_code_from_party_name():
    with pytest.raises(ValueError,match='PARTY_CODE'):
        p.resolve_standard_mapping('Hold',['PARTY_CODE'],['Party Name'],{'Hold':{'PARTY_CODE':'Party Name'}})

def test_same_raw_header_cannot_fill_two_fields():
    m,rows=p.resolve_standard_mapping('Offer',['CREATED_DATETIME','JANGAD_DATE'],['Created Date'],{'Offer':{'JANGAD_DATE':'Created Date'}})
    assert 'CREATED_DATETIME' in m and 'JANGAD_DATE' not in m
    assert next(r for r in rows if r['Final Column']=='JANGAD_DATE')['Status']=='OPTIONAL_MISSING'

def test_phantom_view_helper_only_row_is_safe(tmp_path):
    w=Workbook();ws=w.active;ws.append(['SEQ_NO','APPOINTMENT_DATE','SERIAL_NO','Date']);ws.append([None,None,None,'00:00:00']);ws.append([1,datetime(2026,10,8),27,None]);history=tmp_path/'View.xlsx';w.save(history)
    audit={'issues':[]}
    records,invalid,_=p.read_history('View',history,['SEQ_NO','APPOINTMENT_DATE','SERIAL_NO'],'APPOINTMENT_DATE',audit)
    assert invalid==0 and len(records)==1
    assert any('Helper-only' in x['Issue Type'] for x in audit['issues'])

def test_ambiguous_party_name_does_not_autofill(tmp_path):
    w=Workbook();ws=w.active;ws.append(['Sold Party','PARTY_CODE']);ws.append(['Alpha','P1']);ws.append(['Alpha','P2']);ws.append(['Beta','P3']);name=tmp_path/'Hold.xlsx';w.save(name)
    lookup=p.verified_party_codes(name)
    assert 'ALPHA' not in lookup and lookup['BETA']=='P3'

def test_preflight_accepts_override(tmp_path):
    w=Workbook();ws=w.active;ws.append(['Happened At']);ws.append([datetime(2026,10,9)]);o=io.BytesIO();w.save(o)
    original=preview_auto_pricing_column_mapping({'Offer':o.getvalue()})
    assert next(x for x in original if x['Final column']=='CREATED_DATETIME')['Status']=='MISSING'
    approved=preview_auto_pricing_column_mapping({'Offer':o.getvalue()},column_overrides={'Offer':{'CREATED_DATETIME':'Happened At'}})
    assert next(x for x in approved if x['Final column']=='CREATED_DATETIME')['Source column']=='Happened At'

def test_generic_date_never_replaces_authoritative_hold_date():
    r, rows = p.resolve_standard_mapping('Hold',['JANGAD_DATE'],['Date'])
    assert r == {}
    assert rows[0]['Status']=='MISSING'


def test_previous_sales_formula_date_reads_cached_value(tmp_path):
    from zipfile import ZipFile,ZIP_DEFLATED
    from openpyxl.utils.datetime import to_excel
    import re
    w=Workbook();ws=w.active
    ws.append(['Serial No','Sale Date','Date'])
    ws.append([101,'=INT(C2)', datetime(2026,10,8,14,25)])
    f=tmp_path/'previous.xlsx';w.save(f)
    # openpyxl only writes formula text. Simulate a previous Excel file whose
    # cached formula result was calculated by Excel.
    with ZipFile(f) as src:
        members={n:src.read(n) for n in src.namelist()}
    sheet='xl/worksheets/sheet1.xml'
    txt=members[sheet].decode()
    txt,n=re.subn(r'(<c r="B2"[^>]*><f>INT\(C2\)</f><v>)(</v>)',
        lambda m:m.group(1)+str(to_excel(datetime(2026,10,8)))+m.group(2),txt)
    assert n == 1
    members[sheet]=txt.encode()
    with ZipFile(f,'w',ZIP_DEFLATED) as z:
        for n,data in members.items():z.writestr(n,data)
    audit={'issues':[]}
    records,invalid,_=p.read_sales_history('Sales 120',f,['Serial No','Sale Date'],audit)
    assert invalid==0 and records[0]['Sale Date']==datetime(2026,10,8)
