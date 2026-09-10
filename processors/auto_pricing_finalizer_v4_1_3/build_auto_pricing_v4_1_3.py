from __future__ import annotations

from openpyxl import load_workbook, Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import NamedStyle
from pathlib import Path
from datetime import datetime, date
from collections import defaultdict
from copy import copy
import argparse, gc, math, re, zipfile

SHEETS = ['Sheet1','Sale_segment','Avg_Rate','Hold','Wishlist','View','Offer','Cart']
BASE_FILES = ['Cart.xlsx','Hold.xlsx','Offer.xlsx','View.xlsx','Wishlist.xlsx','Stock file.xlsx','sales_120_file.xlsx','sales_180_file.xlsx']
SPACE_RE = re.compile(r' +')


def clean_text(v) -> str:
    if v is None: return ''
    s = str(v).replace('\u00a0',' ').replace('\t',' ').replace('\n',' ').replace('\r',' ').replace('\u200b','')
    return SPACE_RE.sub(' ', s).strip()


def non(v):
    s = clean_text(v)
    return 'NON' if not s else s


def norm_serial(v) -> str:
    if v is None: return ''
    if isinstance(v, bool): return str(v)
    if isinstance(v, int): return str(v)
    if isinstance(v, float) and math.isfinite(v) and v.is_integer(): return str(int(v))
    s = clean_text(v)
    try:
        f = float(s)
        if math.isfinite(f) and f.is_integer(): return str(int(f))
    except Exception:
        pass
    return s


def tag(shp, grp, col, qua, lab, flo) -> str:
    return '-'.join(clean_text(x) for x in (shp, grp, col, qua, non(lab), non(flo)))


def date_only(v):
    if isinstance(v, datetime): return datetime(v.year,v.month,v.day)
    if isinstance(v, date): return datetime(v.year,v.month,v.day)
    return v


def as_float(v, default=0.0):
    try:
        if v is None or clean_text(v) == '': return default
        return float(v)
    except Exception:
        return default


def read_xlsx_rows(path: Path):
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    it = ws.iter_rows(values_only=True)
    headers = list(next(it))
    ix = {h:i for i,h in enumerate(headers)}
    return wb, it, headers, ix


def capture_template_format(template: Path):
    wb = load_workbook(template, read_only=False, data_only=False)
    fmt = {}; style_reps = {}
    for s in SHEETS:
        if s not in wb.sheetnames:
            raise KeyError(f'Template is missing required sheet: {s}')
        ws = wb[s]
        headers = [ws.cell(1,c).value for c in range(1,ws.max_column+1)]
        hstyles = [ws.cell(1,c).style_id for c in range(1,ws.max_column+1)]
        dstyles = [ws.cell(2,c).style_id for c in range(1,ws.max_column+1)]
        # Date helper style in View can be represented by a later populated row.
        if s == 'View' and ws.max_column >= 26 and ws.cell(3,26).style_id:
            dstyles[25] = ws.cell(3,26).style_id
        for r in (1,2,3):
            for c in range(1,ws.max_column+1):
                cell=ws.cell(r,c); style_reps.setdefault(cell.style_id,cell)
        fmt[s] = {
            'headers':headers, 'hstyles':hstyles, 'dstyles':dstyles,
            'widths':{k:v.width for k,v in ws.column_dimensions.items() if v.width is not None},
            'row1h':ws.row_dimensions[1].height, 'freeze':ws.freeze_panes,
            'grid':ws.sheet_view.showGridLines, 'zoom':ws.sheet_view.zoomScale,
            'zoomn':ws.sheet_view.zoomScaleNormal, 'filter':ws.auto_filter.ref,
        }
    style_defs={}
    for sid,cell in style_reps.items():
        style_defs[sid]={'font':copy(cell.font),'fill':copy(cell.fill),'border':copy(cell.border),
                         'alignment':copy(cell.alignment),'number_format':cell.number_format,
                         'protection':copy(cell.protection)}
    wb.close(); gc.collect()
    return fmt, style_defs


def build_maps(base_dir: Path):
    # 120-day weighted sold average by Tag.
    avg_sum_amt=defaultdict(float); avg_sum_carat=defaultdict(float)
    wb,it,h,ix=read_xlsx_rows(base_dir/'sales_120_file.xlsx')
    for row in it:
        t=tag(row[ix['Shp']],row[ix['Carat_Grp']],row[ix['Col']],row[ix['Qua']],row[ix['Lab']],row[ix['Flo']])
        avg_sum_amt[t]+=as_float(row[ix['Amt']])
        avg_sum_carat[t]+=as_float(row[ix['Carat']])
    wb.close()
    avg_rate={t:(avg_sum_amt[t]/c if c else 0.0) for t,c in avg_sum_carat.items()}

    # 180-day sale segment by Tag: count + ordered unique buyers.
    sale_count=defaultdict(int); buyers=defaultdict(list); buyer_seen=defaultdict(set)
    wb,it,h,ix=read_xlsx_rows(base_dir/'sales_180_file.xlsx')
    for row in it:
        t=tag(row[ix['Shp']],row[ix['Carat_Grp']],row[ix['Col']],row[ix['Qua']],row[ix['Lab']],row[ix['Flo']])
        sale_count[t]+=1
        p=clean_text(row[ix['Sold Party']])
        if p and p not in buyer_seen[t]: buyer_seen[t].add(p); buyers[t].append(p)
    wb.close()

    # Action parties by exact stone serial across all action datasets.
    action_parties=defaultdict(list); action_seen=defaultdict(set)
    for fn in ['Hold.xlsx','Wishlist.xlsx','View.xlsx','Offer.xlsx','Cart.xlsx']:
        wb,it,h,ix=read_xlsx_rows(base_dir/fn)
        serial_col='SERIAL_NO' if 'SERIAL_NO' in ix else 'Serial No'
        for row in it:
            sn=norm_serial(row[ix[serial_col]])
            p=clean_text(row[ix['Sold Party']]) if 'Sold Party' in ix else ''
            if sn and p and p not in action_seen[sn]:
                action_seen[sn].add(p); action_parties[sn].append(p)
        wb.close()
    return avg_rate, sale_count, buyers, action_parties


def build_workbook(template: Path, base_dir: Path, out_path: Path, base_zip: Path|None=None):
    fmt, style_defs = capture_template_format(template)
    avg_rate, sale_count, buyers, action_parties = build_maps(base_dir)

    wb = Workbook(write_only=True)
    style_names={}
    for sid,d in sorted(style_defs.items()):
        name=f'tmpl_{sid}'; ns=NamedStyle(name=name)
        ns.font=d['font']; ns.fill=d['fill']; ns.border=d['border']; ns.alignment=d['alignment']
        ns.number_format=d['number_format']; ns.protection=d['protection']; wb.add_named_style(ns); style_names[sid]=name

    def styled_row(ws, vals, styles):
        cells=[]
        for i,v in enumerate(vals):
            c=WriteOnlyCell(ws,value=v); sid=styles[i] if i<len(styles) else 0
            if sid in style_names: c.style=style_names[sid]
            cells.append(c)
        ws.append(cells)

    def setup(s):
        ws=wb.create_sheet(s); f=fmt[s]
        for col,w in f['widths'].items(): ws.column_dimensions[col].width=w
        if f['row1h'] is not None: ws.row_dimensions[1].height=f['row1h']
        ws.freeze_panes=f['freeze']; ws.sheet_view.showGridLines=f['grid']
        if f['zoom'] is not None: ws.sheet_view.zoomScale=f['zoom']
        if f['zoomn'] is not None: ws.sheet_view.zoomScaleNormal=f['zoomn']
        styled_row(ws,f['headers'],f['hstyles']); return ws

    counts={}
    # Sheet1 fully rebuilt from current Stock + current processed Sales/action bases.
    s='Sheet1'; ws=setup(s); target=fmt[s]['headers']
    swb,sit,sh,six=read_xlsx_rows(base_dir/'Stock file.xlsx'); n=0
    for row in sit:
        rec={h:row[i] for h,i in six.items()}
        sn=norm_serial(rec.get('Serial No')); t=clean_text(rec.get('Tag')) or tag(rec.get('Shp'),rec.get('Carat_Grp'),rec.get('Color name'),rec.get('Qua'),rec.get('Lab'),rec.get('Flo'))
        sold_avg=avg_rate.get(t,0.0); exp=as_float(rec.get('Expected'))
        diff=round((exp-sold_avg)/exp*100,2) if exp else 0.0
        b=buyers.get(t,[]); acts=action_parties.get(sn,[]); actset=set(acts)
        values={
            'Serial No':rec.get('Serial No'),'Shp':rec.get('Shp'),'Carat':rec.get('Carat'),'Carat_Grp':rec.get('Carat_Grp'),
            'Color name':rec.get('Color name'),'Qua':rec.get('Qua'),'Cut':rec.get('Cut'),'Pol':rec.get('Pol'),'Sym':rec.get('Sym'),
            'Lab':non(rec.get('Lab')),'Flo':non(rec.get('Flo')),'GRap':rec.get('GRap'),'GPer':rec.get('GPer'),'Expected':rec.get('Expected'),
            'Sold_Avg_Rate_120days':sold_avg,'Diff':diff,'Remark':('Yes' if sold_avg>0 and diff>0 else 'No'),'Tag':t,
            'Unique_Buyer_Count':len(b),'Sold_Party':(', '.join(b) if b else None),'sale_count':sale_count.get(t,0),
            'Action_count':sum(1 for p in b if p in actset),'ALL_Action_count':len(acts),
        }
        styled_row(ws,[values.get(h) for h in target],fmt[s]['dstyles']); n+=1
    swb.close(); counts[s]=n

    # Sale_segment + Avg_Rate helper sheets, values (not formulas).
    for s,fn in [('Sale_segment','sales_180_file.xlsx'),('Avg_Rate','sales_120_file.xlsx')]:
        ws=setup(s); target=fmt[s]['headers']; swb,it,h,ix=read_xlsx_rows(base_dir/fn); n=0
        for row in it:
            rec={hh:row[ii] for hh,ii in ix.items()}; lab=non(rec.get('Lab')); flo=non(rec.get('Flo'))
            t=tag(rec.get('Shp'),rec.get('Carat_Grp'),rec.get('Col'),rec.get('Qua'),lab,flo); sd=rec.get('Sale Date')
            vals={**rec,'Lab':lab,'Flo':flo,'Date':date_only(sd),'Tag':t}
            styled_row(ws,[vals.get(hh) for hh in target],fmt[s]['dstyles']); n+=1
        swb.close(); counts[s]=n

    # Hold.
    s='Hold'; ws=setup(s); target=fmt[s]['headers']; swb,it,h,ix=read_xlsx_rows(base_dir/'Hold.xlsx'); n=0
    for row in it:
        rec={hh:row[ii] for hh,ii in ix.items()}; rec['Date']=date_only(rec.get('JANGAD_DATE'))
        styled_row(ws,[rec.get(hh) for hh in target],fmt[s]['dstyles']); n+=1
    swb.close(); counts[s]=n

    # Wishlist / Offer / Cart.
    for s in ['Wishlist','Offer','Cart']:
        ws=setup(s); target=fmt[s]['headers']; swb,it,h,ix=read_xlsx_rows(base_dir/(s+'.xlsx')); n=0
        for row in it:
            rec={hh:row[ii] for hh,ii in ix.items()}
            if 'Lab' in rec: rec['Lab']=non(rec['Lab'])
            if 'Flo' in rec: rec['Flo']=non(rec['Flo'])
            styled_row(ws,[rec.get(hh) for hh in target],fmt[s]['dstyles']); n+=1
        swb.close(); counts[s]=n

    # View: one RAW row = one SERIAL_NO value, no comma split; blank/zero already removed by processor.
    s='View'; ws=setup(s); target=fmt[s]['headers']; swb,it,h,ix=read_xlsx_rows(base_dir/'View.xlsx'); n=0
    for row in it:
        rec={hh:row[ii] for hh,ii in ix.items()}; rec['Date']=date_only(rec.get('CREATED_DATETIME'))
        styled_row(ws,[rec.get(hh) for hh in target],fmt[s]['dstyles']); n+=1
    swb.close(); counts[s]=n

    out_path.parent.mkdir(parents=True,exist_ok=True); wb.save(out_path)
    if base_zip:
        with zipfile.ZipFile(base_zip,'w',zipfile.ZIP_DEFLATED) as z:
            for fn in BASE_FILES: z.write(base_dir/fn,arcname=fn)
    return counts


def main():
    ap=argparse.ArgumentParser(description='Build formatted Auto Pricing workbook from v4.1.3 processed base files.')
    ap.add_argument('--template',required=True,help='Existing Auto Pricing workbook used only for sheet layout/styles.')
    ap.add_argument('--base-dir',required=True,help='Directory containing the 8 processed v4.1.3 base xlsx files.')
    ap.add_argument('--output',required=True,help='Final formatted Auto Pricing xlsx path.')
    ap.add_argument('--base-zip',help='Optional clean ZIP containing the 8 base files.')
    a=ap.parse_args()
    counts=build_workbook(Path(a.template),Path(a.base_dir),Path(a.output),Path(a.base_zip) if a.base_zip else None)
    print('Built',a.output); print(counts)

if __name__=='__main__': main()
