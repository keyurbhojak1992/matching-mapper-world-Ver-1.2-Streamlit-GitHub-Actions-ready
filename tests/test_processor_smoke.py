from __future__ import annotations

import io
import zipfile

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

from adapters.bid_drop_adapter import run_bid_drop_checking
from adapters.filter_demand_adapter import run_filter_demand
from adapters.sort_adapter import run_sorting_process
from adapters.suggested_adapter import run_suggested_process


def stock_bytes() -> bytes:
    stock = pd.DataFrame([{
        'Serial No': '1001', 'Shp': 'RBC', 'Carat': 1.05, 'Color': 'D', 'Qua': 'VS1',
        'Flo': 'NON', 'Lab': 'GIA', 'Cut': 'EX', 'Pol': 'EX', 'Sym': 'EX',
        'Shade': 'NONE', 'Milky': 'NONE', 'TBINC_NAME': 'NONE', 'SBINC_NAME': 'NONE',
        'TOIN_NAME': 'NONE', 'SOIN_NAME': 'NONE', 'GIRDLE_OPEN_NAME': 'NONE',
        'PAV_OPEN_NAME': 'NONE', 'Disc%': -35, 'Rate': 5000, 'Amount': 5250,
        'SGPComment': '', 'DETAIL': 'https://example.com/1001'
    }])
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        stock.to_excel(writer, index=False, sheet_name='Sheet1')
    return buffer.getvalue()


def test_suggested_master_adapter_smoke():
    profile = pd.DataFrame([{
        'Salesperson': 'Mr. Test User', 'Sold Party Original': 'TEST PARTY',
        'Normalized Party': 'TEST PARTY',
        'Super Tag': 'RBC|1.00-1.19|D|VS1|NON|GIA|EX|EX|EX|NONE|NONE|NONE|NONE|NONE|NONE|NONE|NONE',
        'Count': 1, 'Match_Logic_Type': 'STANDARD', 'SoldParty': 'TEST PARTY',
        'Country': 'USA', 'City': 'NY', 'BidSaleCount': 1, 'MTD_Label': 'JULY--26',
        'Types_Found': 'SALE', 'Stone_Id': '1001'
    }])
    profile_buffer = io.BytesIO()
    profile.to_parquet(profile_buffer, index=False)

    filters_buffer = io.BytesIO()
    with pd.ExcelWriter(filters_buffer, engine='openpyxl') as writer:
        pd.DataFrame(columns=['Sold Party', 'Sales_Person']).to_excel(writer, index=False, sheet_name='Change Format')

    output, _, _ = run_suggested_process(
        stock_bytes=stock_bytes(),
        profile_perm_bytes=profile_buffer.getvalue(),
        profile_temp_bytes=None,
        filter_bytes=filters_buffer.getvalue(),
        clean_bytes=None,
        runtime_overrides={
            'TEST_MODE': False,
            'TEST_MODE_PARTY': False,
            'ENABLE_CLEAN_FILE': False,
            'GENERATE_MERGED_MASTER': True,
            'GENERATE_INDIVIDUAL_REPORTS': False,
        },
    )
    with zipfile.ZipFile(io.BytesIO(output)) as zf:
        assert 'MASTER_CONSOLIDATED_REPORT.xlsx' in zf.namelist()


def test_filter_demand_adapter_smoke():
    filters_buffer = io.BytesIO()
    with pd.ExcelWriter(filters_buffer, engine='openpyxl') as writer:
        pd.DataFrame([{
            'SOLD PARTY': 'TEST PARTY', 'SALES_PERSON': 'Mr. Test User', 'SHP': 'RBC',
            'CARAT': '1.00 TO 1.10', 'COLOR NAME': 'D', 'QUA': 'VS1', 'FLO': 'NON',
            'LAB': 'GIA', 'REMARKS': 'INCLUDE'
        }]).to_excel(writer, index=False, sheet_name='Change Format')
        pd.DataFrame(columns=['Sold Party']).to_excel(writer, index=False, sheet_name='Removed Party')

    output, _, _ = run_filter_demand(
        stock_bytes=stock_bytes(),
        filter_bytes=filters_buffer.getvalue(),
        runtime_overrides={'TOLERANCE_LEVEL': 1},
    )
    with zipfile.ZipFile(io.BytesIO(output)) as zf:
        assert 'Mr Test User.xlsx' in zf.namelist()


def test_sorting_master_adapter_smoke():
    wb = Workbook()
    nav = wb.active
    nav.title = 'Navigation'
    nav.append(['Go To Sheet', 'Party Name', 'Count'])
    nav.append(['TEST PARTY', 'TEST PARTY', 2])
    ws = wb.create_sheet('TEST PARTY')
    headers = ['Sr.No', 'Serial No', 'Shp', 'Carat', 'Color', 'Qua', 'Cut', 'Pol', 'Sym', 'Flo', 'Lab', 'Disc%', 'SGPComment']
    ws.append(headers)
    ws.append([1, '1001', 'RBC', 1.05, 'D', 'VS1', 'EX', 'EX', 'EX', 'NON', 'GIA', -35, ''])
    green = PatternFill(start_color='CCFFCC', end_color='CCFFCC', fill_type='solid')
    for cell in ws[2][:5]:
        cell.fill = green

    workbook_buffer = io.BytesIO()
    wb.save(workbook_buffer)
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('Mr Test User.xlsx', workbook_buffer.getvalue())

    output, _, _ = run_sorting_process(
        input_zip_bytes=zip_buffer.getvalue(),
        sales_file_bytes=None,
        runtime_overrides={
            'ONLY_MATCHED_TYPE': True,
            'ALLOWED_MATCH_TYPES': {'Exact'},
            'COUNT_THRESHOLD': 0,
        },
    )
    with zipfile.ZipFile(io.BytesIO(output)) as zf:
        assert 'Mr Test User.xlsx' in zf.namelist()


def test_suggested_scope_controls_smoke():
    profile = pd.DataFrame([{
        'Salesperson': 'Mr. Test User', 'Sold Party Original': 'TEST PARTY',
        'Normalized Party': 'TEST PARTY',
        'Super Tag': 'RBC|1.00-1.19|D|VS1|NON|GIA|EX|EX|EX|NONE|NONE|NONE|NONE|NONE|NONE|NONE|NONE',
        'Count': 1, 'Match_Logic_Type': 'STANDARD', 'SoldParty': 'TEST PARTY',
        'Country': 'USA', 'City': 'NY', 'BidSaleCount': 1, 'MTD_Label': 'JULY--26',
        'Types_Found': 'SALE', 'Stone_Id': '1001'
    }])
    profile_buffer = io.BytesIO()
    profile.to_parquet(profile_buffer, index=False)

    filters_buffer = io.BytesIO()
    with pd.ExcelWriter(filters_buffer, engine='openpyxl') as writer:
        pd.DataFrame(columns=['Sold Party', 'Sales_Person']).to_excel(writer, index=False, sheet_name='Change Format')

    output, _, _ = run_suggested_process(
        stock_bytes=stock_bytes(),
        profile_perm_bytes=profile_buffer.getvalue(),
        profile_temp_bytes=None,
        filter_bytes=filters_buffer.getvalue(),
        clean_bytes=None,
        runtime_overrides={
            'TEST_MODE': True,
            'TEST_SALESPEOPLE': ['Mr. Test User'],
            'TEST_MODE_PARTY': True,
            'TEST_PARTIES': ['TEST PARTY'],
            'ENABLE_CLEAN_FILE': False,
            'GENERATE_MERGED_MASTER': True,
            'GENERATE_INDIVIDUAL_REPORTS': False,
            'SHOW_BASIC_SUPERTAG': True,
            'SHOW_ADVANCED_SUPERTAG': False,
            'SHOW_REF_STONE_ID': True,
            'SHOW_CONCAT_STATUS': True,
            'MATCH_STATUS_FILTER_MODE': 'NONE',
            'ENABLE_TYPE_COLUMNS': True,
            'STATIC_TYPE_SEQUENCE': ['Type_SALE', 'Type_BID'],
        },
    )
    with zipfile.ZipFile(io.BytesIO(output)) as zf:
        assert 'MASTER_CONSOLIDATED_REPORT.xlsx' in zf.namelist()



def test_sorting_preserves_rebound_hyperlinks_and_translated_formulas():
    wb = Workbook()
    nav = wb.active
    nav.title = 'Navigation'
    nav.append(['Go To Sheet', 'Count'])
    nav.append(['TEST PARTY', 3])
    ws = wb.create_sheet('TEST PARTY')
    headers = [
        'Sr.No', 'Serial No', 'Shp', 'Carat', 'Color', 'Qua', 'Cut', 'Pol',
        'Sym', 'Flo', 'Lab', 'Disc%', 'SGPComment', 'DETAIL', 'FORMULA LINK'
    ]
    ws.append(headers)
    ws.append([1, 'DROP', 'RBC', 1.05, 'D', 'VS1', 'EX', 'EX', 'EX', 'NON', 'GIA', -35, '', '', ''])
    ws.append([2, 'KEEP', 'RBC', 1.05, 'D', 'VS1', 'EX', 'EX', 'EX', 'NON', 'GIA', -35, '', 'Open', '=HYPERLINK("https://example.com/"&B3,B3)'])
    ws['N3'].hyperlink = 'https://example.com/KEEP'
    green = PatternFill(start_color='CCFFCC', end_color='CCFFCC', fill_type='solid')
    for cell in ws[3][:5]:
        cell.fill = green

    workbook_buffer = io.BytesIO()
    wb.save(workbook_buffer)
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('Test.xlsx', workbook_buffer.getvalue())

    output, _, _ = run_sorting_process(
        input_zip_bytes=zip_buffer.getvalue(),
        sales_file_bytes=None,
        runtime_overrides={
            'ONLY_MATCHED_TYPE': True,
            'ALLOWED_MATCH_TYPES': {'Exact'},
            'COUNT_THRESHOLD': 0,
        },
    )
    with zipfile.ZipFile(io.BytesIO(output)) as zf:
        sorted_wb = load_workbook(io.BytesIO(zf.read('Test.xlsx')), data_only=False)
    sorted_ws = sorted_wb['TEST PARTY']
    assert sorted_ws['B2'].value == 'KEEP'
    assert sorted_ws['N2'].hyperlink.target == 'https://example.com/KEEP'
    assert sorted_ws['N2'].hyperlink.ref == 'N2'
    assert sorted_ws['O2'].value == '=HYPERLINK("https://example.com/"&B2,B2)'


def test_bid_drop_checking_adapter_smoke(tmp_path, monkeypatch):
    import adapters.common as common
    suggested = pd.DataFrame([{'Serial No': '1001', 'Sold Party': 'TEST PARTY'}])
    bids = pd.DataFrame([{'Serial No': '1001', 'Party Name': 'TEST PARTY', 'Qua': 'VS1'}])
    sales = pd.DataFrame([{
        'Sold Party': 'TEST PARTY', 'Shp': 'RBC', 'Carat_grp': '1.00-1.19',
        'Col': 'D', 'Qua': 'VS1', 'Flo': 'NON', 'Lab': 'GIA',
        'Sales_Person': 'MR. TEST USER', 'Country': 'USA'
    }])
    stock = pd.DataFrame([{
        'Serial No': '1001', 'Shp': 'RBC', 'Carat': 1.05, 'Carat_grp': '1.00-1.19',
        'Color': 'D', 'Qua': 'VS1', 'Flo': 'NON', 'Lab': 'GIA', 'Cut': 'EX',
        'Pol': 'EX', 'Sym': 'EX', 'Shade': 'NONE', 'Milky': 'NONE',
        'TBINC_NAME': 'NONE', 'SBINC_NAME': 'NONE', 'Ratio': 1.0
    }])

    def excel_bytes(df, sheet_name='Sheet1'):
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
        return buffer.getvalue()

    filter_buffer = io.BytesIO()
    with pd.ExcelWriter(filter_buffer, engine='openpyxl') as writer:
        pd.DataFrame(columns=['SOLD PARTY', 'SALES_PERSON']).to_excel(
            writer, index=False, sheet_name='Change Format'
        )
        pd.DataFrame(columns=['SOLD PARTY']).to_excel(
            writer, index=False, sheet_name='Removed Party'
        )
        pd.DataFrame(columns=['SOLD PARTY']).to_excel(
            writer, index=False, sheet_name='Tolerance'
        )

    default_dir = tmp_path / 'data'
    default_dir.mkdir()
    (default_dir / 'filters - EBID.xlsx').write_bytes(filter_buffer.getvalue())
    (default_dir / 'no_need.xlsx').write_bytes(
        excel_bytes(pd.DataFrame(columns=['SOLD PARTY']))
    )
    monkeypatch.setattr(common, 'DATA_DIR', default_dir)

    output, filename, _ = run_bid_drop_checking(
        suggested_bytes=excel_bytes(suggested),
        bids_bytes=excel_bytes(bids),
        sales_bytes=excel_bytes(sales),
        stock_bytes=excel_bytes(stock),
        filter_bytes=None,
        no_need_bytes=None,
    )
    assert filename.endswith('.xlsx')
    result = pd.read_excel(io.BytesIO(output))
    assert result.loc[0, 'Diagnostic Status'] == '✅ SUCCESS'
    assert 'suggested and converted' in result.loc[0, 'Explanation']
