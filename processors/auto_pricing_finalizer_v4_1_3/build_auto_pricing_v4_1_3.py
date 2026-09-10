from __future__ import annotations

import argparse
import math
import re
import zipfile
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Border, Font, NamedStyle, PatternFill, Side


BASE_FILES = [
    'Cart.xlsx', 'Hold.xlsx', 'Offer.xlsx', 'View.xlsx', 'Wishlist.xlsx',
    'Stock file.xlsx', 'sales_120_file.xlsx', 'sales_180_file.xlsx'
]
SPACE_RE = re.compile(r' +')

SHEET_HEADERS = {
    'Sheet1': ['Serial No', 'Shp', 'Carat', 'Carat_Grp', 'Color name', 'Qua', 'Cut', 'Pol', 'Sym', 'Lab', 'Flo', 'GRap', 'GPer', 'Expected', 'Sold_Avg_Rate_120days', 'Diff', 'Remark', 'Tag', 'Unique_Buyer_Count', 'Sold_Party', 'sale_count', 'Action_count', 'ALL_Action_count'],
    'Sale_segment': ['Serial No', 'Shp', 'Carat', 'Carat_Grp', 'Col', 'Qua', 'Cut', 'Pol', 'Sym', 'Lab', 'Flo', 'Sold Party', 'Sale Date', 'Date', 'Tag'],
    'Avg_Rate': ['Serial No', 'Shp', 'Carat', 'Carat_Grp', 'Col', 'Qua', 'Cut', 'Pol', 'Sym', 'Lab', 'Flo', 'Sale Date', 'SPer', 'SRate', 'Amt', 'Date', 'Tag'],
    'Hold': ['SERIAL_NO', 'JANGAD_DATE', 'Date', 'PARTY_CODE', 'Sold Party'],
    'Wishlist': ['CREATED_DATETIME', 'Sold Party', 'Party_type', 'Country', 'City', 'SALE_PERSON', 'Serial No', 'Shp', 'Carat', 'Col', 'Qua', 'Flo', 'Lab', 'Cut', 'Pol', 'Sym', 'Shade', 'Milky', 'TBINC_NAME', 'TOIN_NAME', 'SBINC_NAME', 'SOIN_NAME', 'GIRDLE_OPEN_NAME', 'PAV_OPEN_NAME', 'Ratio'],
    'View': ['SEQ_NO', 'APPOINTMENT_TYPE', 'APPOINTMENT_DATE', 'APPOINTMENT_TIME', 'PARTY_CODE', 'Sold Party', 'EMPLOYEE_NO', 'SERIAL_NO', 'SOURCE', 'STATUS', 'CREATED_BY', 'CREATED_DATETIME', 'CREATED_COMPUTER', 'MODIFIED_BY', 'MODIFIED_DATETIME', 'MODIFIED_COMPUTER', 'IS_REACHED', 'IS_CANCELED', 'NOTE', 'ASSIGN_TO', 'STOCK_REQUEST_TIME', 'APPOINTMENT_END_TIME', 'IS_DELETE', 'IS_LAYOUT', 'CABIN_ASSIGN_DATE', 'Date'],
    'Offer': ['CREATED_DATETIME', 'Sold Party', 'Party_type', 'Country', 'City', 'SALE_PERSON', 'Serial No', 'Shp', 'Carat', 'Col', 'Qua', 'Flo', 'Lab', 'Cut', 'Pol', 'Sym', 'Shade', 'Milky', 'TBINC_NAME', 'TOIN_NAME', 'SBINC_NAME', 'SOIN_NAME', 'GIRDLE_OPEN_NAME', 'PAV_OPEN_NAME', 'Ratio'],
    'Cart': ['CREATED_DATETIME', 'Sold Party', 'Party_type', 'Country', 'City', 'SALE_PERSON', 'Serial No', 'Shp', 'Carat', 'Col', 'Qua', 'Flo', 'Lab', 'Cut', 'Pol', 'Sym', 'Shade', 'Milky', 'TBINC_NAME', 'TOIN_NAME', 'SBINC_NAME', 'SOIN_NAME', 'GIRDLE_OPEN_NAME', 'PAV_OPEN_NAME', 'Ratio'],
}

SHEET_WIDTHS = {
    'Sheet1': {'A':9,'B':6.3,'C':5.6,'D':9.9,'E':11.2,'F':5.6,'G':4.6,'H':3.9,'I':4.8,'J':5.3,'K':5.3,'L':6,'M':7.8,'N':10,'O':22.6,'P':7.8,'Q':8,'R':33.6,'S':20.3,'T':30.8,'U':10.6,'V':12.8,'W':16.9},
    'Sale_segment': {'A':9.6,'B':5.9,'C':7,'D':10.3,'E':5,'F':5,'G':5,'H':5,'I':5,'J':5,'K':5,'L':44.9,'M':18,'N':15.6,'O':33.6},
    'Avg_Rate': {'A':9.6,'B':5.9,'C':7,'D':10.3,'E':5,'F':5,'G':5,'H':5,'I':5,'J':5,'K':5,'L':18,'M':8,'N':10,'O':10.5,'P':15.6,'Q':33.2},
    'Hold': {'A':10.8,'B':22,'C':10.5,'D':12.5,'E':57.3},
    'Wishlist': {'A':19.2,'B':45.3,'C':10.6,'D':17,'E':13.5,'F':13.5,'G':9,'H':6.3,'I':6,'J':4.8,'K':5.6,'L':5.3,'M':6.3,'N':4,'O':3.9,'P':4.8,'Q':22.3,'R':12.5,'S':18.2,'T':14.2,'U':18.2,'V':14.2,'W':20.2,'X':17.6,'Y':5.6},
    'Offer': {'A':19.2,'B':45.6,'C':10.6,'D':17,'E':15.8,'F':13.5,'G':9,'H':6.3,'I':6,'J':4.8,'K':5.6,'L':5.3,'M':6.3,'N':4,'O':3.9,'P':4.8,'Q':22.3,'R':12.5,'S':18.2,'T':14.2,'U':18.2,'V':14.2,'W':20.2,'X':17.6,'Y':5.6},
    'Cart': {'A':19.2,'B':41.6,'C':10.6,'D':17,'E':15.6,'F':13.5,'G':9,'H':6.3,'I':6,'J':4.8,'K':5.6,'L':5.3,'M':6.3,'N':4,'O':3.9,'P':4.8,'Q':22.3,'R':12.5,'S':18.2,'T':14.2,'U':18.2,'V':14.2,'W':20.2,'X':17.6,'Y':5.6},
    'View': {'A':8.3,'B':19.9,'C':20.2,'D':20.3,'E':12.6,'F':45.6,'G':14.5,'H':10.8,'I':9,'J':10.3,'K':12.2,'L':19.2,'M':20.5,'N':13.5,'O':20.3,'P':21.6,'Q':11.8,'R':12.8,'S':56.5,'T':11.2,'U':21.6,'V':24.9,'W':9.8,'X':10.8,'Y':20.2,'Z':18.3},
}


def clean_text(v) -> str:
    if v is None:
        return ''
    s = str(v).replace('\u00a0', ' ').replace('\t', ' ').replace('\n', ' ').replace('\r', ' ').replace('\u200b', '')
    return SPACE_RE.sub(' ', s).strip()


def non(v):
    s = clean_text(v)
    return 'NON' if not s else s


def norm_serial(v) -> str:
    if v is None:
        return ''
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, int):
        return str(v)
    if isinstance(v, float) and math.isfinite(v) and v.is_integer():
        return str(int(v))
    s = clean_text(v)
    try:
        f = float(s)
        if math.isfinite(f) and f.is_integer():
            return str(int(f))
    except Exception:
        pass
    return s


def tag(shp, grp, col, qua, lab, flo) -> str:
    return '-'.join(clean_text(x) for x in (shp, grp, col, qua, non(lab), non(flo)))


def date_only(v):
    if isinstance(v, datetime):
        return datetime(v.year, v.month, v.day)
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    return v


def as_float(v, default=0.0):
    try:
        if v is None or clean_text(v) == '':
            return default
        return float(v)
    except Exception:
        return default


def read_xlsx_rows(path: Path):
    """Read the first worksheet and normalize every data row to header width.

    Excel does not physically store trailing blank cells. In read-only mode openpyxl
    can therefore return a shorter tuple for rows whose right-most fields are blank.
    The Auto Pricing builder addresses columns by their header position, so pad those
    rows with ``None`` instead of treating a valid blank cell as a malformed row.
    """
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    source_rows = ws.iter_rows(values_only=True)
    try:
        headers = list(next(source_rows))
    except StopIteration as exc:
        wb.close()
        raise ValueError(f"Workbook has no header row: {path.name}") from exc

    width = len(headers)
    ix = {h: i for i, h in enumerate(headers) if h is not None}

    def padded_rows():
        for row in source_rows:
            row = tuple(row)
            if len(row) < width:
                row = row + (None,) * (width - len(row))
            yield row

    return wb, padded_rows(), headers, ix


def build_maps(base_dir: Path):
    avg_sum_amt = defaultdict(float)
    avg_sum_carat = defaultdict(float)
    wb, it, _, ix = read_xlsx_rows(base_dir / 'sales_120_file.xlsx')
    for row in it:
        t = tag(row[ix['Shp']], row[ix['Carat_Grp']], row[ix['Col']], row[ix['Qua']], row[ix['Lab']], row[ix['Flo']])
        avg_sum_amt[t] += as_float(row[ix['Amt']])
        avg_sum_carat[t] += as_float(row[ix['Carat']])
    wb.close()
    avg_rate = {t: (avg_sum_amt[t] / c if c else 0.0) for t, c in avg_sum_carat.items()}

    sale_count = defaultdict(int)
    buyers = defaultdict(list)
    buyer_seen = defaultdict(set)
    wb, it, _, ix = read_xlsx_rows(base_dir / 'sales_180_file.xlsx')
    for row in it:
        t = tag(row[ix['Shp']], row[ix['Carat_Grp']], row[ix['Col']], row[ix['Qua']], row[ix['Lab']], row[ix['Flo']])
        sale_count[t] += 1
        p = clean_text(row[ix['Sold Party']])
        if p and p not in buyer_seen[t]:
            buyer_seen[t].add(p)
            buyers[t].append(p)
    wb.close()

    action_parties = defaultdict(list)
    action_seen = defaultdict(set)
    for fn in ['Hold.xlsx', 'Wishlist.xlsx', 'View.xlsx', 'Offer.xlsx', 'Cart.xlsx']:
        wb, it, _, ix = read_xlsx_rows(base_dir / fn)
        serial_col = 'SERIAL_NO' if 'SERIAL_NO' in ix else 'Serial No'
        for row in it:
            sn = norm_serial(row[ix[serial_col]])
            p = clean_text(row[ix['Sold Party']]) if 'Sold Party' in ix else ''
            if sn and p and p not in action_seen[sn]:
                action_seen[sn].add(p)
                action_parties[sn].append(p)
        wb.close()
    return avg_rate, sale_count, buyers, action_parties


def _register_styles(wb: Workbook):
    thin = Side(style='thin', color='000000')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def add(name, *, font=None, fill=None, alignment=None, number_format='General', border_value=None):
        style = NamedStyle(name=name)
        if font is not None:
            style.font = font
        if fill is not None:
            style.fill = fill
        if alignment is not None:
            style.alignment = alignment
        style.number_format = number_format
        if border_value is not None:
            style.border = border_value
        wb.add_named_style(style)

    add('main_header', font=Font(name='Calibri', size=11, bold=True), alignment=Alignment(horizontal='center', vertical='top'), border_value=border)
    add('main_data', font=Font(name='Calibri', size=11), alignment=Alignment(horizontal='center'), border_value=border)
    add('helper_header', font=Font(name='Microsoft Sans Serif', size=9, bold=True), alignment=Alignment(horizontal='center', vertical='center'), border_value=border)
    add('helper_center', font=Font(name='Microsoft Sans Serif', size=8), alignment=Alignment(horizontal='center', vertical='center'), border_value=border)
    add('helper_right', font=Font(name='Microsoft Sans Serif', size=8), alignment=Alignment(horizontal='right', vertical='center'), border_value=border)
    add('helper_left', font=Font(name='Microsoft Sans Serif', size=8), alignment=Alignment(horizontal='left', vertical='center'), border_value=border)
    add('helper_datetime', font=Font(name='Microsoft Sans Serif', size=8), alignment=Alignment(horizontal='center', vertical='center'), number_format='m/d/yy h:mm', border_value=border)
    add('helper_date', font=Font(name='Calibri', size=11), alignment=Alignment(horizontal='center'), number_format='mm-dd-yy', border_value=border)
    add('action_header', font=Font(name='Calibri', size=11, bold=True), alignment=Alignment(horizontal='center', vertical='center'), border_value=border)
    add('action_data', font=Font(name='Calibri', size=11), alignment=Alignment(horizontal='center', vertical='center'), border_value=border)
    add('action_date', font=Font(name='Calibri', size=11), alignment=Alignment(horizontal='center', vertical='center'), number_format='mm-dd-yy', border_value=border)
    add('action_serial', font=Font(name='Calibri', size=11, bold=True), fill=PatternFill('solid', fgColor='FCE4D6'), alignment=Alignment(horizontal='center', vertical='center'), border_value=border)
    add('view_header', font=Font(name='Calibri', size=11, bold=True), alignment=Alignment(horizontal='center', vertical='top'), border_value=border)
    add('view_date_header', font=Font(name='Calibri', size=11, bold=True), fill=PatternFill('solid', fgColor='FFFF00'), alignment=Alignment(horizontal='center', vertical='top'), border_value=border)
    add('view_datetime', font=Font(name='Calibri', size=11), number_format='yyyy-mm-dd hh:mm:ss')
    add('hold_datetime', font=Font(name='Calibri', size=11), number_format='yyyy-mm-dd hh:mm:ss.000')
    add('hold_date', font=Font(name='Calibri', size=11), number_format='mm-dd-yy')


def _style_sequence(sheet: str, headers: list[str], *, header=False):
    if sheet == 'Sheet1':
        return ['main_header' if header else 'main_data'] * len(headers)
    if sheet in {'Sale_segment', 'Avg_Rate'}:
        if header:
            return ['helper_header'] * len(headers)
        styles = []
        for h in headers:
            if h == 'Sold Party':
                styles.append('helper_left')
            elif h == 'Sale Date':
                styles.append('helper_datetime')
            elif h == 'Date':
                styles.append('helper_date')
            elif h in {'Carat', 'SPer'}:
                styles.append('helper_right')
            elif h == 'Tag':
                styles.append('main_data')
            else:
                styles.append('helper_center')
        return styles
    if sheet == 'Hold':
        if header:
            return [None] * len(headers)
        return [None, 'hold_datetime', 'hold_date', None, None]
    if sheet == 'View':
        if header:
            return ['view_date_header' if h == 'Date' else 'view_header' for h in headers]
        date_cols = {'APPOINTMENT_DATE', 'CREATED_DATETIME', 'MODIFIED_DATETIME', 'CABIN_ASSIGN_DATE', 'Date'}
        return ['view_datetime' if h in date_cols else None for h in headers]
    if sheet in {'Wishlist', 'Offer', 'Cart'}:
        if header:
            return ['action_header'] * len(headers)
        styles = []
        for h in headers:
            if h == 'CREATED_DATETIME':
                styles.append('action_date')
            elif h == 'Serial No':
                styles.append('action_serial')
            else:
                styles.append('action_data')
        return styles
    return [None] * len(headers)


def build_workbook(base_dir: Path, out_path: Path, base_zip: Path | None = None):
    avg_rate, sale_count, buyers, action_parties = build_maps(base_dir)
    wb = Workbook(write_only=True)
    _register_styles(wb)

    def styled_row(ws, vals, styles):
        cells = []
        for value, style in zip(vals, styles):
            cell = WriteOnlyCell(ws, value=value)
            if style:
                cell.style = style
            cells.append(cell)
        ws.append(cells)

    def setup(sheet):
        ws = wb.create_sheet(sheet)
        headers = SHEET_HEADERS[sheet]
        for col, width in SHEET_WIDTHS.get(sheet, {}).items():
            ws.column_dimensions[col].width = width
        styled_row(ws, headers, _style_sequence(sheet, headers, header=True))
        return ws, headers

    counts = {}

    ws, target = setup('Sheet1')
    swb, sit, _, six = read_xlsx_rows(base_dir / 'Stock file.xlsx')
    n = 0
    for row in sit:
        rec = {h: row[i] for h, i in six.items()}
        sn = norm_serial(rec.get('Serial No'))
        t = clean_text(rec.get('Tag')) or tag(rec.get('Shp'), rec.get('Carat_Grp'), rec.get('Color name'), rec.get('Qua'), rec.get('Lab'), rec.get('Flo'))
        sold_avg = avg_rate.get(t, 0.0)
        exp = as_float(rec.get('Expected'))
        diff = round((exp - sold_avg) / exp * 100, 2) if exp else 0.0
        b = buyers.get(t, [])
        acts = action_parties.get(sn, [])
        actset = set(acts)
        values = {
            'Serial No': rec.get('Serial No'), 'Shp': rec.get('Shp'), 'Carat': rec.get('Carat'), 'Carat_Grp': rec.get('Carat_Grp'),
            'Color name': rec.get('Color name'), 'Qua': rec.get('Qua'), 'Cut': rec.get('Cut'), 'Pol': rec.get('Pol'), 'Sym': rec.get('Sym'),
            'Lab': non(rec.get('Lab')), 'Flo': non(rec.get('Flo')), 'GRap': rec.get('GRap'), 'GPer': rec.get('GPer'), 'Expected': rec.get('Expected'),
            'Sold_Avg_Rate_120days': sold_avg, 'Diff': diff, 'Remark': ('Yes' if sold_avg > 0 and diff > 0 else 'No'), 'Tag': t,
            'Unique_Buyer_Count': len(b), 'Sold_Party': (', '.join(b) if b else None), 'sale_count': sale_count.get(t, 0),
            'Action_count': sum(1 for p in b if p in actset), 'ALL_Action_count': len(acts),
        }
        styled_row(ws, [values.get(h) for h in target], _style_sequence('Sheet1', target))
        n += 1
    swb.close()
    counts['Sheet1'] = n

    for sheet, filename in [('Sale_segment', 'sales_180_file.xlsx'), ('Avg_Rate', 'sales_120_file.xlsx')]:
        ws, target = setup(sheet)
        swb, it, _, ix = read_xlsx_rows(base_dir / filename)
        n = 0
        styles = _style_sequence(sheet, target)
        for row in it:
            rec = {h: row[i] for h, i in ix.items()}
            lab = non(rec.get('Lab'))
            flo = non(rec.get('Flo'))
            t = tag(rec.get('Shp'), rec.get('Carat_Grp'), rec.get('Col'), rec.get('Qua'), lab, flo)
            vals = {**rec, 'Lab': lab, 'Flo': flo, 'Date': date_only(rec.get('Sale Date')), 'Tag': t}
            styled_row(ws, [vals.get(h) for h in target], styles)
            n += 1
        swb.close()
        counts[sheet] = n

    ws, target = setup('Hold')
    swb, it, _, ix = read_xlsx_rows(base_dir / 'Hold.xlsx')
    n = 0
    styles = _style_sequence('Hold', target)
    for row in it:
        rec = {h: row[i] for h, i in ix.items()}
        rec['Date'] = date_only(rec.get('JANGAD_DATE'))
        styled_row(ws, [rec.get(h) for h in target], styles)
        n += 1
    swb.close()
    counts['Hold'] = n

    for sheet in ['Wishlist', 'Offer', 'Cart']:
        ws, target = setup(sheet)
        swb, it, _, ix = read_xlsx_rows(base_dir / f'{sheet}.xlsx')
        n = 0
        styles = _style_sequence(sheet, target)
        for row in it:
            rec = {h: row[i] for h, i in ix.items()}
            if 'Lab' in rec:
                rec['Lab'] = non(rec['Lab'])
            if 'Flo' in rec:
                rec['Flo'] = non(rec['Flo'])
            styled_row(ws, [rec.get(h) for h in target], styles)
            n += 1
        swb.close()
        counts[sheet] = n

    ws, target = setup('View')
    swb, it, _, ix = read_xlsx_rows(base_dir / 'View.xlsx')
    n = 0
    styles = _style_sequence('View', target)
    for row in it:
        rec = {h: row[i] for h, i in ix.items()}
        rec['Date'] = date_only(rec.get('CREATED_DATETIME'))
        styled_row(ws, [rec.get(h) for h in target], styles)
        n += 1
    swb.close()
    counts['View'] = n

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)

    if base_zip:
        with zipfile.ZipFile(base_zip, 'w', zipfile.ZIP_DEFLATED) as zf:
            for filename in BASE_FILES:
                zf.write(base_dir / filename, arcname=filename)
    return counts


def main():
    ap = argparse.ArgumentParser(description='Build the formatted Auto Pricing workbook from prepared files.')
    ap.add_argument('--base-dir', required=True)
    ap.add_argument('--output', required=True)
    ap.add_argument('--base-zip')
    args = ap.parse_args()
    counts = build_workbook(Path(args.base_dir), Path(args.output), Path(args.base_zip) if args.base_zip else None)
    print('Built', args.output)
    print(counts)


if __name__ == '__main__':
    main()
