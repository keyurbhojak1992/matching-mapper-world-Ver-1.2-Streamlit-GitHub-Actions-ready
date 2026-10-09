from __future__ import annotations

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from pathlib import Path
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
from collections import Counter, defaultdict
import math
import re
import json
import zipfile
import os
from difflib import SequenceMatcher
from openpyxl.styles.numbers import is_date_format

VERSION = '4.1.9-optional-pricing-fields'
RAW_DIR = Path(os.environ.get('RAW_DATA_DIR', '/mnt/data/raw_inputs'))
FINAL_DIR = Path(os.environ.get('FINAL_HISTORY_DIR', '/mnt/data/final_samples'))
# Do not create output directories when imported for a read-only mapping preview.
# Streamlit Cloud may not allow writes to the legacy /mnt/data default path.
OUT_DIR = Path(os.environ.get('OUTPUT_DIR', '/mnt/data/output_final'))

ACTIVITY_DAYS = 90
HOLD_DAYS = 90
SALES_WINDOWS = [(120, 'sales_120_file.xlsx'), (180, 'sales_180_file.xlsx')]
STANDARD_SALES_HEADERS = ['Serial No', 'Shp', 'Carat', 'Carat_Grp', 'Col', 'Qua', 'Cut', 'Pol', 'Sym', 'Lab', 'Flo', 'Sold Party', 'Sale Date', 'SPer', 'SRate', 'Amt']

BLACKLIST_TEXT = r'''
SHUBHAMT
HRITHIK
AMIT
AAKASHJ
RUSHI
SONALI
NIKUNJVS
DHAVAL
SUNIL
PARVA
AATISH
KALASH
HIREN
SHUBHAMP
AAKASH
ARUN
VIVEKCHL
ADMIN
BHAVIK
PRANAV
MUKUL
JAGDISH
NEELAY
RIDHAM
PRIYAL
TEJ
HARSHAD
ZEAL
JATIN
LAXMI
SMITD
SHUBHAM
PRIYANK
NIKUNJN
SONALI   SONANAVE
Shubham   Tarsariya
Hrithik Hundia
Amit Tarsaria
AMIT TARSARIYA
Aakash Jain
Rushi Shankar
Sonali   Sonavane
Nikunj Shankar
Dhaval   Moradiya
Sunil Bhatiya
Parva Shah
Aatish   Sarodiya
Kalash   Sanghavi
Hiren Rathod
Shubham Patil
Aakash Parmar
Arun Dholiya
Vivek   Chalodiya
Bhavik Hadiya
Pranav   Chalodiya
Mukul Shah
Jagdish   Tarsariya
Neelay Sampat
Ridham Lakhani
Priyal Shankar
Tej Shankar
Harshad Pawar
Zeal Harsora
Jatin   Motisariya
Laxmi Rathod
Smit Dayani
Priyank   Kothari
Nikunj   Vastrapara
NIKUNJ   VASTARPARA
HET SHANKAR
NIKUNJVS   SHANKAR
RAVI   KASODARIYA
SHIVAM JEWELS   TEMP USER
SONALI   SONANAVE
SHIVAM JEWELS   INC
QUALITY CHECK
TEMPORARY
SHIVAM JEWELS   FZCO
S JEWELS HK   LTD.
SHIVAM
SHOW
Mr. Shubham   Tarsariya
Mr. Hrithik   Hundia
Mr. Amit   Tarsaria
Mr. Aakash   Jain
Mr. Rushi   Shankar
Ms. Sonali   Sonavane
Mr. Nikunj   Shankar
Mr. Dhaval   Moradiya
Mr. Sunil   Bhatiya
Mr. Parva Shah
Mr. Aatish   Sarodiya
Mr. Kalash   Sanghavi
Mr. Hiren   Rathod
Mr. Shubham   Patil
Mr. Aakash   Parmar
Mr. Arun   Dholiya
Mr. Vivek   Chalodiya
Mr. Bhavik   Hadiya
Mr. Pranav   Chalodiya
Mr. Mukul Shah
Mr. Jagdish   Tarsariya
Mr. Neelay   Sampat
Mr. Ridham   Lakhani
Mr. Priyal   Shankar
Mr. Tej   Shankar
Mr. Harshad   Pawar
Ms. Zeal   Harsora
Mr. Jatin   Motisariya
Ms. Laxmi   Rathod
Mr. Smit   Dayani
Mr. Priyank   Kothari
Mr. Nikunj   Vastrapara
Mr. HET   SHANKAR
Mr. RAVI   KASODARIYA
Shivam
Ms. SONALI   SONANAVE
MANUAL VIDEO   REQUEST
LOOSE CHECK
ADMIN PATEL
PRIYAL SHANKAR
ASHWIN BHAI   SHANKAR
'''

CONTROL_RE = re.compile(r'[\x00-\x1f]')
SPACE_RE = re.compile(r' +')
DERIVED_COLUMNS = {'Carat_Grp', 'Amt', 'Tag'}
# Only missing columns that affect the builder's calculations or safe record-window
# processing may block. Other fields stay in the stable output schema as blanks,
# are shown in the audit, and can still be explicitly mapped by the user.
AUTO_PRICING_REQUIRED_COLUMNS = {
    'Cart': {'CREATED_DATETIME', 'Serial No', 'Sold Party'},
    'Offer': {'CREATED_DATETIME', 'Serial No', 'Sold Party'},
    'Wishlist': {'CREATED_DATETIME', 'Serial No', 'Sold Party'},
    'Hold': {'JANGAD_DATE', 'SERIAL_NO', 'Sold Party'},
    'View': {'APPOINTMENT_DATE', 'SERIAL_NO', 'Sold Party'},
    # Pricing tag, weighted average, sales count and buyer metrics depend on these.
    'Sales': {'Serial No', 'Shp', 'Carat', 'Col', 'Qua', 'Lab', 'Flo',
              'Sold Party', 'Sale Date', 'SRate'},
    # GRap and GPer are only mandatory when Expected must be calculated.
    'Stock': {'Serial No', 'Shp', 'Carat', 'Color name', 'Qua', 'Lab', 'Flo', 'Expected'},
}


def pricing_required_column(dataset: str, target: str, resolved=None) -> bool:
    if target in DERIVED_COLUMNS or norm_header(target) in LEGACY_FINAL_COLUMNS:
        return False
    if dataset == 'Stock' and target in {'GRap', 'GPer'}:
        return 'Expected' not in (resolved or {})
    return target in AUTO_PRICING_REQUIRED_COLUMNS.get(dataset, set())


def missing_mapping_rows(rows):
    return [r for r in rows if r['Status'] == 'MISSING']


def record_mapping_issues(audit, dataset, rows, raw_path):
    for r in rows:
        if r['Status'] not in {'MISSING', 'OPTIONAL_MISSING'}:
            continue
        blocking = r['Status'] == 'MISSING'
        add_issue(audit, dataset, 'ERROR' if blocking else 'WARNING',
                  'Missing required pricing/source column' if blocking else 'Optional source column missing',
                  source_label(raw_path), column=r['Final Column'],
                  details=(('Processing must not proceed without this business-date, record key or pricing input. ')
                           if blocking else ('Not used by pricing calculations, activity matching or processing windows. '
                                               'New source records retain a blank for this output field. '))
                          + f"Expected headers: {r['Candidate Raw Columns']}")
STOCK_FINAL_HEADERS = ['Serial No', 'Shp', 'Carat', 'Carat_Grp', 'Color name', 'Qua', 'Cut', 'Pol', 'Sym', 'Lab', 'Flo', 'GRap', 'GPer', 'Expected', 'Tag']
LEGACY_FINAL_COLUMNS = {'date'}  # Obsolete helper column; never used for current-date/window logic.
MAPPING_FILE = Path(os.environ.get('COLUMN_MAPPING_FILE', str(Path(__file__).with_name('column_mapping_aliases.json'))))
OVERRIDE_FILE = os.environ.get('COLUMN_MAPPING_OVERRIDE_FILE', '')
PARTY_CODE_MASTER_FILE = os.environ.get('PARTY_CODE_MASTER_FILE', '')

ALIASES = {
    'Hold': {
        'SERIAL_NO': ['SERIAL_NO', 'Serial No', 'STONE_ID'],
        'JANGAD_DATE': ['JANGAD_DATE', 'Jangad Date', 'Hold Date'],
        'PARTY_CODE': ['PARTY_CODE', 'Party Code'],
        'Sold Party': ['Sold Party', 'PARTY_NAME', 'Party Name'],
    },
    'View': {
        'SEQ_NO': ['SEQ_NO'], 'APPOINTMENT_TYPE': ['APPOINTMENT_TYPE'],
        'APPOINTMENT_DATE': ['APPOINTMENT_DATE'], 'APPOINTMENT_TIME': ['APPOINTMENT_TIME'],
        'PARTY_CODE': ['PARTY_CODE'], 'Sold Party': ['Sold Party', 'PARTY_NAME', 'Party Name'],
        'EMPLOYEE_NO': ['EMPLOYEE_NO'], 'SERIAL_NO': ['SERIAL_NO', 'STONE_ID', 'Serial No'],
        'SOURCE': ['SOURCE'], 'STATUS': ['STATUS'], 'CREATED_BY': ['CREATED_BY'],
        'CREATED_DATETIME': ['CREATED_DATETIME'], 'CREATED_COMPUTER': ['CREATED_COMPUTER'],
        'MODIFIED_BY': ['MODIFIED_BY'], 'MODIFIED_DATETIME': ['MODIFIED_DATETIME'],
        'MODIFIED_COMPUTER': ['MODIFIED_COMPUTER'], 'IS_REACHED': ['IS_REACHED'],
        'IS_CANCELED': ['IS_CANCELED'], 'NOTE': ['NOTE'], 'ASSIGN_TO': ['ASSIGN_TO'],
        'STOCK_REQUEST_TIME': ['STOCK_REQUEST_TIME'], 'APPOINTMENT_END_TIME': ['APPOINTMENT_END_TIME'],
        'IS_DELETE': ['IS_DELETE'], 'IS_LAYOUT': ['IS_LAYOUT'], 'CABIN_ASSIGN_DATE': ['CABIN_ASSIGN_DATE'],
    },
    'Sales': {
        # Canonical final names stay fixed even when future RAW exports rename headers.
        'Serial No': ['Serial No', 'Serial Number', 'SerialNumber', 'Sr No', 'Sr. No.', 'SR_NO', 'NO.', 'NO', 'SERIAL_NO', 'SERIALNO', 'STONE_ID', 'STONE NO', 'STONE_NO', 'Packet No', 'PACKET_NO', 'LOT_NO'],
        'Shp': ['Shp', 'Shape', 'Shape Name', 'SHAPE_NAME', 'SHAPE_CODE', 'S_CODE'],
        'Carat': ['Carat', 'CTS', 'CT', 'Weight', 'WEIGHT', 'WT', 'CARAT', 'CARAT_WEIGHT'],
        'Col': ['Col', 'Color', 'Colour', 'COLOR', 'COLOUR', 'Color name', 'Colour Name', 'GCOLOR_NAME', 'G_COLOR'],
        'Qua': ['Qua', 'Quality', 'Clarity', 'CLARITY', 'Clarity Name', 'GCLARITY_NAME', 'G_CLARITY'],
        'Cut': ['Cut', 'Cut Name', 'CUT_NAME', 'GCUT_NAME', 'G_CUT'],
        'Pol': ['Pol', 'Polish', 'Polish Name', 'POLISH_NAME', 'GPOLISH_NAME', 'G_POLISH'],
        'Sym': ['Sym', 'Symmetry', 'Symmetry Name', 'SYMMENTRY_NAME', 'SYMMETRY', 'GSYM_NAME', 'G_SYM'],
        'Lab': ['Lab', 'Laboratory', 'LAB_NAME', 'LABORATORY'],
        'Flo': ['Flo', 'FL', 'Fluorescence', 'Fluor', 'FL_NAME', 'GFL_NAME', 'G_FL'],
        'Sold Party': ['Sold Party', 'Sold To', 'Buyer', 'Customer', 'Customer Name', 'Client', 'Party', 'PARTY_NAME', 'Party Name', 'pary name'],
        'Sale Date': ['Sale Date', 'Sales Date', 'Sold Date', 'Date Sold', 'S date', 'S_DATE', 'SDATE', 'SALE_DATE', 'SALE_DT', 'SALES_DATE', 'SOLD_DATE'],
        'SPer': ['SPer', 'SALE_PER', 'Sale %', 'Sale Percent', 'Discount %', 'Discount', 'S.%', 'S%', 'SALEPER', 'Sale Per'],
        'SRate': ['SRate', 'Rate', 'Sale Rate', 'Sales Rate', 'SALE_RATE', 'S RATE', 'SRATE', 'PRICE_RATE'],
    },
    'Stock': {
        'Serial No': ['Serial No', 'SERIAL_NO', 'NO.', 'STONE_ID'], 'Shp': ['Shp', 'Shape'],
        'Carat': ['Carat', 'CTS', 'WEIGHT'], 'Color name': ['Color name', 'Color', 'Col'],
        'Qua': ['Qua', 'Quality', 'Clarity'], 'Cut': ['Cut'], 'Pol': ['Pol'], 'Sym': ['Sym'],
        'Lab': ['Lab'], 'Flo': ['Flo', 'FL'], 'GRap': ['GRap', 'RAP'],
        'GPer': ['GPer', 'GPER', 'G.%'], 'Expected': ['Expected', 'Exp rate', 'Expected Rate', 'EXP_RATE', 'EXP AMT'],
    }
}


def processing_date() -> date:
    """Production: actual Asia/Kolkata date. Regression only: TEST_PROCESSING_DATE=YYYY-MM-DD."""
    test_value = clean_text(os.environ.get('TEST_PROCESSING_DATE', ''))
    if test_value:
        try:
            return datetime.strptime(test_value, '%Y-%m-%d').date()
        except ValueError as exc:
            raise ValueError('TEST_PROCESSING_DATE must be YYYY-MM-DD') from exc
    return datetime.now(ZoneInfo('Asia/Kolkata')).date()


def source_name_map() -> dict:
    path = RAW_DIR / '_source_names.json'
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding='utf-8'))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def source_label(path: Path) -> str:
    return source_name_map().get(path.name, path.name)


def clean_text(v: str) -> str:
    s = str(v)
    s = s.replace('\u00a0', ' ').replace('\t', ' ').replace('\n', ' ').replace('\r', ' ').replace('\u200b', '')
    s = CONTROL_RE.sub('', s)
    s = SPACE_RE.sub(' ', s).strip()
    return s


def is_blank(v) -> bool:
    if v is None:
        return True
    if isinstance(v, str):
        return clean_text(v) == ''
    return False


def clean_value(v):
    return clean_text(v) if isinstance(v, str) else v


def norm_party(v) -> str:
    return clean_text('' if v is None else str(v)).upper()


def load_blacklist() -> set[str]:
    """Use Remove party.xlsx when supplied; otherwise fall back to the validated embedded list."""
    path_text = clean_text(os.environ.get('REMOVE_PARTY_FILE', ''))
    if path_text:
        path = Path(path_text)
        if not path.exists():
            raise FileNotFoundError(f'REMOVE_PARTY_FILE not found: {path}')
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        header = [clean_text(v or '') for v in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
        idx = None
        for candidate in ('Sales_Short_name', 'Sold Party', 'Party Name', 'Party'):
            if candidate in header:
                idx = header.index(candidate)
                break
        if idx is None:
            idx = 1 if len(header) > 1 else 0
        values = set()
        for row in ws.iter_rows(min_row=2, values_only=True):
            if idx < len(row) and not is_blank(row[idx]):
                values.add(norm_party(row[idx]))
        wb.close()
        if not values:
            raise ValueError(f'No removal-party values found in {path}')
        return values
    return {norm_party(x) for x in BLACKLIST_TEXT.splitlines() if clean_text(x)}


BLACKLIST = load_blacklist()


def norm_header(h) -> str:
    if h is None:
        return ''
    return re.sub(r'[^a-z0-9]+', '', clean_text(str(h)).lower())


def parse_datetime(v):
    if v is None or v == '':
        return None
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if 20000 <= float(v) <= 80000:
            from openpyxl.utils.datetime import from_excel
            try: return from_excel(v)
            except Exception: return None
        return None
    s = clean_text(str(v))
    if not s or '#REF!' in s.upper() or s.startswith('='):
        return None
    for fmt in (
        '%Y-%m-%d %H:%M:%S.%f','%Y-%m-%d %H:%M:%S','%Y-%m-%d',
        '%d-%m-%Y %H:%M:%S.%f','%d-%m-%Y %H:%M:%S','%d-%m-%Y',
        '%m/%d/%Y %H:%M:%S','%m/%d/%Y','%d/%m/%Y %H:%M:%S','%d/%m/%Y'
    ):
        try: return datetime.strptime(s, fmt)
        except ValueError: pass
    try:
        from dateutil.parser import parse
        return parse(s)
    except Exception:
        return None


def serial_normalize(v):
    v = clean_value(v)
    if v is None or isinstance(v, bool): return v
    if isinstance(v, int): return v
    if isinstance(v, float):
        if math.isfinite(v) and v.is_integer(): return int(v)
        return v
    s = clean_text(str(v))
    if re.fullmatch(r'\d+', s):
        try: return int(s)
        except Exception: return s
    if re.fullmatch(r'\d+\.0+', s):
        try: return int(float(s))
        except Exception: return s
    return s


def number_normalize(v):
    if v is None or v == '': return v
    if isinstance(v, (int, float)) and not isinstance(v, bool): return v
    s = clean_text(str(v))
    try:
        if re.fullmatch(r'[-+]?\d+', s): return int(s)
        if re.fullmatch(r'[-+]?(?:\d+\.\d*|\d*\.\d+)', s): return float(s)
    except Exception: pass
    return clean_value(v)


def read_headers(path: Path):
    wb = load_workbook(path, read_only=True, data_only=False)
    ws = wb[wb.sheetnames[0]]
    vals = list(next(ws.iter_rows(min_row=1, max_row=1, values_only=True)))
    while vals and vals[-1] is None: vals.pop()
    wb.close()
    return [clean_text(str(x)) if x is not None else '' for x in vals]


def read_final_headers(path: Path):
    # Previous finals can contain an old helper column named `Date` (often =INT(...)).
    # It was used by the former manual-date workflow and is no longer part of the
    # required output schema or raw-column validation. Business date fields such as
    # CREATED_DATETIME, JANGAD_DATE and Sale Date remain authoritative.
    return [h for h in read_headers(path) if norm_header(h) not in LEGACY_FINAL_COLUMNS]


def rows_as_dicts(path: Path, data_only: bool = False):
    wb = load_workbook(path, read_only=True, data_only=data_only)
    ws = wb[wb.sheetnames[0]]
    it = ws.iter_rows(values_only=True)
    header_row = next(it)
    headers = [clean_text(str(x)) if x is not None else '' for x in header_row]
    last_nonempty = max((i for i, h in enumerate(headers) if h), default=-1)
    headers = headers[:last_nonempty + 1]
    try:
        for source_row, row in enumerate(it, start=2):
            row = row[:len(headers)]
            if not any(not is_blank(v) for v in row): continue
            yield source_row, {headers[i]: row[i] if i < len(row) else None for i in range(len(headers))}
    finally:
        wb.close()


def _external_aliases():
    try:
        data = json.loads(MAPPING_FILE.read_text(encoding='utf-8')) if MAPPING_FILE.exists() else {}
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def verified_mapping_overrides() -> dict:
    """Only a user-approved, per-run JSON mapping may resolve unknown headers."""
    if not OVERRIDE_FILE:
        return {}
    data = json.loads(Path(OVERRIDE_FILE).read_text(encoding='utf-8'))
    if not isinstance(data, dict) or any(not isinstance(v, dict) for v in data.values()):
        raise ValueError('Column mapping overrides must be {dataset: {final_column: raw_header}}')
    return data


def _safe_override(dataset: str, target: str, raw: str):
    if dataset == 'Hold' and target == 'PARTY_CODE' and norm_header(raw) in {'partyname', 'soldparty', 'paryname'}:
        raise ValueError('PARTY_CODE cannot be taken from Party Name; provide a verified party-code source')
    # Different business dates must not be conflated by an unreviewed heuristic.
    if (target in {'CREATED_DATETIME', 'JANGAD_DATE', 'APPOINTMENT_DATE', 'Sale Date'}
        and norm_header(raw) in {'date', 'createddatetime'} and target in {'JANGAD_DATE', 'APPOINTMENT_DATE', 'Sale Date'}):
        raise ValueError(f'{dataset}.{target}: {raw!r} is not an authoritative business date')


def resolve_standard_mapping(dataset: str, final_headers: list[str], raw_headers: list[str], overrides=None):
    """Exact approved alias -> explicit override -> suggestion only; one-to-one source mapping."""
    overrides = verified_mapping_overrides().get(dataset, {}) if overrides is None else overrides.get(dataset, {})
    if not isinstance(overrides, dict):
        raise ValueError('Mappings must be a dictionary per dataset')
    unknown_targets = set(overrides) - set(final_headers)
    if unknown_targets:
        raise ValueError(f'{dataset}: mapping references unknown final columns {sorted(unknown_targets)}')
    raw_headers = [clean_text(str(h)) for h in raw_headers if h and clean_text(str(h))]
    normalized = defaultdict(list)
    for h in raw_headers: normalized[norm_header(h)].append(h)
    used = set()
    resolved = {}
    rows = []
    for target in final_headers:
        aliases = candidates_for(dataset, target)
        if target in DERIVED_COLUMNS:
            rows.append({'Dataset':dataset, 'Final Column':target, 'Candidate Raw Columns':'', 'Raw Columns Found':'',
                         'Status':'DERIVED', 'Rule':'Calculated in output', 'Mapping Method':'DERIVED', 'Confidence':1.0})
            continue
        selected = None
        method = 'EXACT_ALIAS'
        if target in overrides:
            chosen = clean_text(str(overrides[target]))
            _safe_override(dataset, target, chosen)
            found = normalized.get(norm_header(chosen), [])
            if len(found) == 1:
                selected = found[0]
                method = 'APPROVED_OVERRIDE'
            else:
                method = 'OVERRIDE_NOT_FOUND'
        else:
            # Canonical target always precedes aliases, regardless of RAW file order.
            for alias in [target] + [a for a in aliases if norm_header(a) != norm_header(target)]:
                found = normalized.get(norm_header(alias), [])
                if found:
                    selected = found[0]
                    if len(found) > 1: method = 'DUPLICATE_HEADER'
                    break
        if selected and (selected in used or len(normalized[norm_header(selected)]) > 1):
            method = 'CONFLICT'
        if selected and method != 'CONFLICT':
            used.add(selected)
            resolved[target] = selected
            rows.append({'Dataset':dataset,'Final Column':target,'Candidate Raw Columns':' | '.join(aliases),
                         'Raw Columns Found':selected,'Status':'OK','Rule':'Approved column mapping',
                         'Mapping Method':method,'Confidence':1.0,'Suggestion':''})
        else:
            scores = []
            for source in raw_headers:
                if source in used: continue
                ns = norm_header(source)
                if len(ns)<4:continue
                score = max((SequenceMatcher(None,ns,norm_header(a)).ratio() for a in aliases if len(norm_header(a))>=4),default=0)
                if score >= .75: scores.append((score, source))
            scores.sort(reverse=True)
            suggested = scores[0][1] if scores and scores[0][0] >= .80 else ''
            status = 'MISSING'
            rows.append({'Dataset':dataset, 'Final Column':target,'Candidate Raw Columns':' | '.join(aliases),
                         'Raw Columns Found':'','Status':status,
                         'Rule':('Duplicate header/source conflict; remove ambiguity' if method=='CONFLICT'
                                 else 'Unknown field: approve a mapping to avoid an incorrect substitution'),
                         'Mapping Method':method if method != 'EXACT_ALIAS' else 'UNMAPPED',
                         'Confidence':round(scores[0][0],3) if scores else 0.0,
                         'Suggestion':suggested})
    # Classify after resolving ALL fields so Stock can derive Expected from GRap/GPer
    # if a direct Expected source column was not provided.
    for row in rows:
        required = pricing_required_column(dataset, row['Final Column'], resolved)
        row['Pricing Impact'] = 'REQUIRED' if required else 'OPTIONAL'
        if row['Status'] == 'MISSING' and not required:
            row['Status'] = 'OPTIONAL_MISSING'
            row['Rule'] = ('Not needed for Auto Pricing calculation/keys/date windows. '
                           'Blank in new output records unless an approved mapping is supplied.')
    return resolved, rows


def candidates_for(dataset: str, target: str):
    built_in = list(ALIASES.get(dataset, {}).get(target, [target]))
    extra = _external_aliases().get(dataset, {}).get(target, [])
    if isinstance(extra, str):
        extra = [extra]
    if not isinstance(extra, list):
        extra = []
    return list(dict.fromkeys(built_in + [clean_text(str(x)) for x in extra if clean_text(str(x))]))


SALES_HEADER_HINTS = {
    'Serial No': ('serial', 'stone', 'packet', 'lotno', 'srno'),
    'Shp': ('shape',),
    'Carat': ('carat', 'weight', 'cts'),
    'Col': ('color', 'colour'),
    'Qua': ('clarity', 'quality'),
    'Cut': ('cut',),
    'Pol': ('polish',),
    'Sym': ('symmetry', 'symmentry'),
    'Lab': ('laboratory', 'lab'),
    'Flo': ('fluorescence', 'fluor'),
    'Sold Party': ('soldparty', 'soldto', 'partyname', 'customer', 'buyer', 'client'),
    'SPer': ('saleper', 'salepercent', 'discountpercent', 'discountpct'),
    'SRate': ('salerate', 'salesrate', 'pricerate', 'srate'),
}


def _string_looks_like_date(value) -> bool:
    if not isinstance(value, str):
        return False
    text = clean_text(value)
    if not text:
        return False
    patterns = (
        r'\b(?:19|20)\d{2}[-/.]\d{1,2}[-/.]\d{1,2}\b',
        r'\b\d{1,2}[-/.]\d{1,2}[-/.](?:\d{2}|\d{4})\b',
        r'\b\d{1,2}[- ](?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*[- ](?:\d{2}|\d{4})\b',
        r'\b(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*[ -]\d{1,2}[, -]+(?:\d{2}|\d{4})\b',
    )
    return any(re.search(p, text, re.I) for p in patterns)


def profile_columns(path: Path, sample_rows: int = 500):
    """Profile RAW columns so unknown Sales date headers can be recognized by values/Excel format."""
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    header_cells = next(ws.iter_rows(min_row=1, max_row=1))
    headers = [clean_text(str(c.value)) if c.value is not None else '' for c in header_cells]
    last = max((i for i, h in enumerate(headers) if h), default=-1)
    headers = headers[:last + 1]
    stats = {h: {'nonblank': 0, 'parseable_date': 0, 'explicit_date': 0, 'date_format': 0} for h in headers if h}
    try:
        for ridx, row in enumerate(ws.iter_rows(min_row=2, max_row=sample_rows + 1), start=2):
            for i, h in enumerate(headers):
                if not h or i >= len(row):
                    continue
                cell = row[i]
                value = cell.value
                if is_blank(value):
                    continue
                st = stats[h]
                st['nonblank'] += 1
                parsed = parse_datetime(value)
                if parsed is not None:
                    st['parseable_date'] += 1
                explicit = isinstance(value, (datetime, date)) or _string_looks_like_date(value)
                try:
                    fmt_date = bool(cell.number_format and is_date_format(cell.number_format))
                except Exception:
                    fmt_date = False
                if fmt_date:
                    st['date_format'] += 1
                if explicit or (fmt_date and parsed is not None):
                    st['explicit_date'] += 1
    finally:
        wb.close()
    for h, st in stats.items():
        n = st['nonblank'] or 1
        st['parse_ratio'] = st['parseable_date'] / n
        st['explicit_ratio'] = st['explicit_date'] / n
        st['format_ratio'] = st['date_format'] / n
    return stats


def _date_header_bonus(header: str) -> float:
    n = norm_header(header)
    bonus = 0.0
    if 'date' in n: bonus += 25.0
    if any(x in n for x in ('saledt', 'salesdt', 'solddt', 'sdate', 'saledate', 'salesdate', 'solddate')): bonus += 35.0
    if any(x in n for x in ('invoice', 'bill', 'txn', 'transaction')) and ('date' in n or n.endswith('dt')): bonus += 20.0
    if n.endswith('dt'): bonus += 8.0
    if any(x in n for x in ('serial', 'stone', 'packet', 'code', 'rate', 'price', 'carat', 'weight', 'percent', 'per')): bonus -= 45.0
    return bonus


def detect_sale_date_column(raw_path: Path, raw_headers: list[str], exclude=None):
    """Choose an unfamiliar Sales date column using actual date values and Excel number formats."""
    exclude = set(exclude or ())
    profiles = profile_columns(raw_path)
    best = None
    for h in raw_headers:
        if not h or h in exclude:
            continue
        st = profiles.get(h, {})
        n = st.get('nonblank', 0)
        if n < 2:
            continue
        score = _date_header_bonus(h)
        score += 55.0 * st.get('explicit_ratio', 0)
        score += 30.0 * st.get('format_ratio', 0)
        score += 15.0 * st.get('parse_ratio', 0)
        # A completely unfamiliar header is still acceptable when the cells themselves are strongly date-like.
        strong_values = st.get('explicit_ratio', 0) >= 0.80 or st.get('format_ratio', 0) >= 0.80
        if not strong_values and score < 55:
            continue
        candidate = (score, st.get('explicit_ratio', 0), st.get('format_ratio', 0), st.get('parse_ratio', 0), h)
        if best is None or candidate > best:
            best = candidate
    if best is None:
        return None, 0.0, 'MISSING'
    score, explicit_ratio, format_ratio, parse_ratio, h = best
    confidence = min(0.99, max(explicit_ratio, format_ratio, parse_ratio) * 0.75 + min(max(score, 0), 100) / 400)
    return h, confidence, 'DATE_PROFILE'


def _best_fuzzy_sales_header(target: str, raw_headers: list[str], used: set[str]):
    aliases = candidates_for('Sales', target)
    best = None
    for raw in raw_headers:
        if not raw or raw in used:
            continue
        nr = norm_header(raw)
        if len(nr) < 4:
            continue
        for alias in aliases:
            na = norm_header(alias)
            if len(na) < 4:
                continue
            sim = SequenceMatcher(None, nr, na).ratio()
            if sim >= 0.86:
                cand = (sim, raw, alias)
                if best is None or cand > best:
                    best = cand
    return best


def resolve_sales_mapping(raw_path: Path, final_headers: list[str] | None = None, overrides=None):
    """Sales uses the same safe resolver as every dataset; no fuzzy auto-assignment."""
    headers = list(final_headers or STANDARD_SALES_HEADERS)
    return resolve_standard_mapping('Sales', headers, read_headers(raw_path), overrides=overrides)


def map_row_with_sources(row: dict, source_map: dict[str, str], final_headers: list[str]):
    out = {}
    used = {}
    for target in final_headers:
        if target in DERIVED_COLUMNS:
            continue
        source = source_map.get(target)
        value = row.get(source) if source else None
        out[target] = clean_value(value) if not is_blank(value) else None
        used[target] = source
    return out, used


def sales_structure_score(path: Path):
    """Score a workbook as a Sales RAW export, allowing arbitrary filenames."""
    try:
        resolved, rows = resolve_sales_mapping(path, STANDARD_SALES_HEADERS)
    except Exception:
        return -1, {}
    key_targets = {'Serial No', 'Carat', 'Sold Party', 'Sale Date', 'SPer', 'SRate'}
    key_found = sum(1 for k in key_targets if k in resolved)
    total_found = sum(1 for h in STANDARD_SALES_HEADERS if h not in DERIVED_COLUMNS and h in resolved)
    # Require a real Sale Date plus enough Sales-specific structure; this prevents Stock/View from being mistaken for Sales.
    if 'Sale Date' not in resolved or key_found < 4 or total_found < 7:
        return -1, resolved
    return key_found * 10 + total_found, resolved


def locate_sales(root: Path):
    files = [p for p in root.glob('*.xlsx') if not p.name.startswith('~$')]
    # Fast path: validate obvious Sales-named files first. This keeps normal runs fast
    # while preserving the structural fallback for arbitrary future filenames.
    named = [p for p in files if re.search(r'(sale|sdate|sold)', p.stem, re.I)]
    pools = [named, files] if named else [files]
    for pool in pools:
        ranked = []
        for p in pool:
            score, _ = sales_structure_score(p)
            if score >= 0:
                name_bonus = 5 if re.search(r'(sale|sdate|sold)', p.stem, re.I) else 0
                ranked.append((score + name_bonus, p.name.lower(), p))
        if ranked:
            ranked.sort(reverse=True)
            return ranked[0][2]
    raise FileNotFoundError(f'No Sales-like workbook found in {root}; filename may vary, but a date-like Sales column and Sales field structure are required.')


def map_row(row: dict, dataset: str, final_headers: list[str], source_map=None):
    if source_map is None:
        source_map, _ = resolve_standard_mapping(dataset, final_headers, list(row))
    return map_row_with_sources(row, source_map, final_headers)


def party_rejected(rec: dict) -> bool:
    for h, v in rec.items():
        if norm_header(h) in {'soldparty', 'partyname', 'paryname'} and norm_party(v) in BLACKLIST:
            return True
    return False


def normalize_lab_flo(rec: dict):
    # Business rule: blank Lab/Flo are represented as NON in final outputs.
    for h in ('Lab', 'Flo'):
        if h in rec and is_blank(rec.get(h)):
            rec[h] = 'NON'
    return rec


def activity_anchor(raw_dates: list[datetime], today: date) -> date:
    if any(dt.date() == today for dt in raw_dates): return today
    return today - timedelta(days=1)


def key_text(rec: dict, fields):
    parts = []
    for f in fields:
        v = rec.get(f)
        if isinstance(v, datetime): v = v.isoformat(sep=' ')
        parts.append(f'{f}={v}')
    return '; '.join(parts)


def add_issue(audit, dataset, severity, issue_type, source_file='', source_row='', key='', column='', raw_value='', details=''):
    audit['issues'].append({
        'Dataset': dataset, 'Severity': severity, 'Issue Type': issue_type, 'Source File': source_file,
        'Source Row': source_row, 'Key': key, 'Column': column, 'Raw Value': raw_value, 'Details': details
    })


def locate(patterns, root: Path):
    files = [p for p in root.glob('*.xlsx') if not p.name.startswith('~$')]
    for pat in patterns:
        rx = re.compile(pat, re.I)
        hits = [p for p in files if rx.search(p.stem)]
        if hits: return sorted(hits)[0]
    raise FileNotFoundError(f'No file matching {patterns} in {root}')


def final_path(name: str):
    p = FINAL_DIR / name
    if p.exists(): return p
    raise FileNotFoundError(f'Missing previous final/history file: {name}')


def schema_mapping(dataset: str, final_headers: list[str], raw_headers: list[str], overrides=None):
    resolved, rows = resolve_standard_mapping(dataset, final_headers, raw_headers, overrides=overrides)
    # Stock Expected uses the existing GRap/GPer verified calculation if not supplied.
    if dataset == 'Stock' and 'Expected' not in resolved and {'GRap','GPer'} <= set(resolved):
        for row in rows:
            if row['Final Column'] == 'Expected' and row['Status'] == 'MISSING':
                row.update(Status='DERIVED', Rule='GRap/GPer calculation with numeric row validation',
                           **{'Mapping Method':'STOCK_FORMULA', 'Suggestion':''})
    return rows


def read_history(dataset, history_path: Path, headers, date_col, audit):
    records = []
    invalid = 0
    blacklist_removed = 0
    samples = 0
    blacklist_samples = 0
    for source_row, row in rows_as_dicts(history_path):
        rec = {h: clean_value(row.get(h)) for h in headers}
        # Previous exports sometimes contain a helper-Date-only phantom row.
        # Never mistake that artifact for a real activity with a missing business date.
        if all(is_blank(v) for v in rec.values()):
            add_issue(audit, dataset, 'INFO', 'Helper-only empty history row ignored', history_path.name, source_row,
                      details='All business fields were empty; obsolete Date helper was not used as an activity date.')
            continue
        normalize_lab_flo(rec)
        if 'Serial No' in rec: rec['Serial No'] = serial_normalize(rec.get('Serial No'))
        if 'SERIAL_NO' in rec: rec['SERIAL_NO'] = serial_normalize(rec.get('SERIAL_NO'))
        dt = parse_datetime(rec.get(date_col))
        if dt is None:
            invalid += 1
            if samples < 50:
                add_issue(audit, dataset, 'ERROR', 'Invalid date in previous final/history', history_path.name, source_row,
                          column=date_col, raw_value=rec.get(date_col),
                          details='Processing for this dataset must not silently drop historical rows. Upload a clean previous final file.')
                samples += 1
            continue
        rec[date_col] = dt
        if party_rejected(rec):
            blacklist_removed += 1
            if blacklist_samples < 25:
                add_issue(audit, dataset, 'INFO', 'Removed party excluded from previous file', history_path.name, source_row,
                          key=key_text(rec, ['Serial No' if 'Serial No' in rec else 'SERIAL_NO', date_col]),
                          details='Party/Sold Party matched the saved removal list after text cleaning.')
                blacklist_samples += 1
            continue
        rec['_source_row'] = source_row
        records.append(rec)
    return records, invalid, blacklist_removed


def _exact_output_signature(rec: dict, headers: list[str]):
    """Signature for true output-row identity. Internal/source metadata is excluded."""
    vals = []
    for h in headers:
        v = rec.get(h)
        if is_blank(v):
            vals.append(('blank', ''))
        elif isinstance(v, datetime):
            vals.append(('datetime', v.isoformat()))
        elif isinstance(v, date):
            vals.append(('date', v.isoformat()))
        else:
            vals.append((type(v).__name__, str(v)))
    return tuple(vals)


def remove_exact_output_duplicates(records: list[dict], headers: list[str]):
    """Remove only rows whose complete output values are identical; preserve first occurrence/order."""
    seen = set()
    kept = []
    removed = 0
    for rec in records:
        sig = _exact_output_signature(rec, headers)
        if sig in seen:
            removed += 1
            continue
        seen.add(sig)
        kept.append(rec)
    return kept, removed


def verified_party_codes(history_path: Path):
    """Build only unambiguous party-name-to-code matches from previous Hold history."""
    matches = defaultdict(set)
    if history_path.is_file():
        for _, row in rows_as_dicts(history_path):
            name = norm_party(row.get('Sold Party'))
            code = clean_text(str(row.get('PARTY_CODE') or ''))
            if name and code:
                matches[name].add(code)
    if PARTY_CODE_MASTER_FILE:
        master = Path(PARTY_CODE_MASTER_FILE)
        for _, row in rows_as_dicts(master):
            normed = {norm_header(k):v for k,v in row.items()}
            name = norm_party(normed.get('partyname') or normed.get('soldparty'))
            code = clean_text(str(normed.get('partycode') or ''))
            if name and code:
                matches[name].add(code)
    return {name: next(iter(codes)) for name,codes in matches.items() if len(codes) == 1}


def process_history_log(dataset: str, raw_path: Path, history_path: Path, out_name: str, date_col: str,
                        days: int, hold_mode=False, audit=None):
    """History + RAW processor with authoritative whole-boundary-date refresh.

    If RAW contains the last calendar date present in the previous final, every previous-final
    row from that date is discarded and that whole date is rebuilt from RAW. This protects
    against partial-day previous finals that ended at a timestamp before the day was complete.
    """
    headers = read_final_headers(history_path)
    raw_headers = read_headers(raw_path)
    source_map, _ = resolve_standard_mapping(dataset, headers, raw_headers)
    mapping_rows = schema_mapping(dataset, headers, raw_headers)
    if dataset == 'Hold' and 'PARTY_CODE' not in source_map and history_path.is_file():
        for m in mapping_rows:
            if m['Final Column'] == 'PARTY_CODE':
                m.update(Status='OPTIONAL_LOOKUP',
                         Rule='Best-effort verified history/master lookup. Unmatched codes remain blank; never block pricing.',
                         **{'Mapping Method':'VERIFIED_PARTY_LOOKUP'})
    audit['mapping'].extend(mapping_rows)
    record_mapping_issues(audit, dataset, mapping_rows, raw_path)

    history, history_invalid, history_blacklist = read_history(dataset, history_path, headers, date_col, audit)
    code_lookup = verified_party_codes(history_path) if dataset == 'Hold' and 'PARTY_CODE' not in source_map else {}
    unresolved_parties = Counter()
    if history_invalid:
        summary = make_failed_summary(dataset, out_name, raw_path, history_path, date_col, f'{history_invalid} invalid history date row(s)')
        audit['summary'].append(summary)
        return summary, False

    history_max = max((r[date_col] for r in history), default=None)
    raw_rows = raw_valid = raw_invalid = blacklist_removed = 0
    raw_dates = []
    valid_raw = []
    invalid_samples = blacklist_samples = 0

    for source_row, row in rows_as_dicts(raw_path):
        raw_rows += 1
        rec, _ = map_row(row, dataset, headers, source_map)
        normalize_lab_flo(rec)
        if 'Serial No' in rec: rec['Serial No'] = serial_normalize(rec.get('Serial No'))
        if 'SERIAL_NO' in rec: rec['SERIAL_NO'] = serial_normalize(rec.get('SERIAL_NO'))
        dt = parse_datetime(rec.get(date_col))
        if dt is None:
            raw_invalid += 1
            if invalid_samples < 25:
                add_issue(audit, dataset, 'ERROR', 'Invalid raw date', source_label(raw_path), source_row, column=date_col,
                          raw_value=rec.get(date_col), details='Raw row excluded because date cannot be parsed.')
                invalid_samples += 1
            continue
        raw_valid += 1
        raw_dates.append(dt)
        rec[date_col] = dt
        if party_rejected(rec):
            blacklist_removed += 1
            if blacklist_samples < 25:
                add_issue(audit, dataset, 'INFO', 'Blacklist row removed', source_label(raw_path), source_row,
                          key=key_text(rec, ['Serial No' if 'Serial No' in rec else 'SERIAL_NO', date_col]),
                          details='Party/Sold Party matched removal list after CLEAN/TRIM.')
                blacklist_samples += 1
            continue
        if dataset == 'Hold' and 'PARTY_CODE' not in source_map:
            name = norm_party(rec.get('Sold Party'))
            code = code_lookup.get(name)
            if code:
                rec['PARTY_CODE'] = code
            else:
                unresolved_parties[name or '(BLANK PARTY NAME)'] += 1
        rec['_source_row'] = source_row
        valid_raw.append(rec)

    boundary_history_removed = 0
    carried_history = history
    appended = []
    raw_old_overlap = 0
    boundary_refreshed = False

    if history_max is None:
        appended = valid_raw
    else:
        boundary_date = history_max.date()
        # General rule for Cart/Hold/Offer/Wishlist and any caller using this routine:
        # if RAW contains the previous final's last date, RAW is authoritative for that whole day.
        has_boundary_raw = any(r[date_col].date() == boundary_date for r in valid_raw)
        if has_boundary_raw:
            boundary_refreshed = True
            boundary_history_removed = sum(1 for r in history if r[date_col].date() == boundary_date)
            carried_history = [r for r in history if r[date_col].date() < boundary_date]
            appended = [r for r in valid_raw if r[date_col].date() >= boundary_date]
            raw_old_overlap = len(valid_raw) - len(appended)
            add_issue(audit, dataset, 'INFO', 'Boundary date refreshed from RAW', source_label(raw_path),
                      details=f'Previous-final rows dated {boundary_date} were removed and that whole date was rebuilt from current RAW; later RAW dates were appended.')
        else:
            appended = [r for r in valid_raw if r[date_col] > history_max]
            raw_old_overlap = len(valid_raw) - len(appended)

    if raw_old_overlap:
        rule_txt = ('earlier than the refreshed boundary date' if boundary_refreshed
                    else f'<= latest previous-final timestamp {history_max}')
        add_issue(audit, dataset, 'INFO', 'Historical raw overlap not appended', source_label(raw_path),
                  details=f'{raw_old_overlap} raw row(s) are {rule_txt}. They were not appended.')

    today = processing_date()
    anchor = today if hold_mode else activity_anchor(raw_dates, today)
    cutoff = anchor - timedelta(days=days)
    if not hold_mode and anchor != today:
        add_issue(audit, dataset, 'INFO', 'No current-day log', source_label(raw_path),
                  details=f'No raw record dated {today}; anchor moved to {anchor}.')

    combined = carried_history + appended
    before = after = 0
    filtered = []
    for rec in combined:
        d = rec[date_col].date()
        if d < cutoff: before += 1
        elif d > anchor: after += 1
        else: filtered.append(rec)
    filtered.sort(key=lambda r: (r[date_col], r.get('_source_row', 0)))

    # Remove only true full-row duplicates after merge/window filtering.
    # Similar-looking activity with any differing output field is preserved.
    filtered, exact_duplicates_removed = remove_exact_output_duplicates(filtered, headers)
    missing_party_codes = Counter(norm_party(r.get('Sold Party')) for r in filtered if dataset=='Hold' and is_blank(r.get('PARTY_CODE')))
    if missing_party_codes:
        total_missing = sum(missing_party_codes.values())
        add_issue(audit, dataset, 'WARNING', 'Optional Hold PARTY_CODE not verified', source_label(raw_path),
                  column='PARTY_CODE',
                  details=f'{total_missing} Hold final row(s) have no verified code across {len(missing_party_codes)} parties. '
                          'Codes remain blank; serial and Sold Party still drive Auto Pricing. A code master is optional.')
        for party, count in missing_party_codes.most_common(10):
            add_issue(audit, dataset, 'INFO', 'Party code lookup not available', source_label(raw_path),
                      column='PARTY_CODE', raw_value=party,
                      details=f'{count} retained Hold final row(s) with blank optional code.')
    if exact_duplicates_removed:
        add_issue(audit, dataset, 'INFO', 'Exact duplicate rows removed',
                  details=f'{exact_duplicates_removed} fully identical output row(s) were removed; the first occurrence was retained.')

    write_standard_output(out_name, headers, filtered)
    audit_blanks(audit, dataset, headers, filtered)
    status = 'PASS' if raw_invalid == 0 and not missing_mapping_rows(mapping_rows) else 'FAILED'
    summary = {
        'Dataset': dataset, 'Output File': out_name, 'Raw File': source_label(raw_path), 'Previous Final File': history_path.name,
        'Source Rule': 'Previous final preserved before its latest calendar date; if RAW contains that date, the whole boundary date is replaced from RAW; later RAW dates appended. Otherwise append only after latest timestamp.',
        'Window Rule': f'{days}-day window; ' + ('anchor=current processing date' if hold_mode else 'anchor=today when current-day log exists, otherwise yesterday'),
        'Date Field': date_col, 'Processing Date': today.isoformat(), 'Anchor Date': anchor.isoformat(), 'Cutoff Date': cutoff.isoformat(),
        'History Rows': len(history), 'History Invalid Dates': history_invalid, 'History Blacklist Removed': history_blacklist, 'Latest History Timestamp': history_max,
        'Raw Rows Read': raw_rows, 'Raw Valid Dates': raw_valid, 'Raw Invalid Dates': raw_invalid,
        'Raw New Rows Appended': len(appended), 'Raw Historical/Overlap Rows Skipped': raw_old_overlap,
        'Boundary History Rows Replaced': boundary_history_removed,
        'Blacklist Removed From Raw': blacklist_removed, 'Before Cutoff': before, 'After Anchor': after,
        'Output Rows': len(filtered), 'Exact Duplicate Rows Removed': exact_duplicates_removed, 'Potential Duplicates Kept': 0,
        'New Duplicate-Looking Occurrences This Run': '',
        'Total Duplicate-Looking Occurrences In Final': '', 'Status': status
    }
    audit['summary'].append(summary)
    return summary, status == 'PASS'

def _split_view_serials(value):
    """Explode a View STONE_ID/SERIAL_NO cell into individual stone serials."""
    if is_blank(value):
        return []
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        one = serial_normalize(value)
        return [] if one in (None, '', 0, '0', '0.0') else [one]
    text = clean_text(str(value))
    if not text:
        return []
    out = []
    for part in re.split(r'[,;\r\n]+', text):
        serial = serial_normalize(part)
        if serial in (None, '', 0, '0', '0.0'):
            continue
        out.append(serial)
    return out


def process_view(raw_path: Path, history_path: Path, out_name='View.xlsx', audit=None):
    dataset = 'View'
    date_col = 'APPOINTMENT_DATE'
    headers = read_final_headers(history_path)
    raw_headers = read_headers(raw_path)
    source_map, _ = resolve_standard_mapping(dataset, headers, raw_headers)
    mapping_rows = schema_mapping(dataset, headers, raw_headers)
    audit['mapping'].extend(mapping_rows)
    record_mapping_issues(audit, dataset, mapping_rows, raw_path)

    # View aging and boundary refresh follow the appointment date, not the row-created timestamp.
    history, history_invalid, history_blacklist = read_history(dataset, history_path, headers, date_col, audit)
    history_clean = []
    history_bad_stone = 0
    for rec in history:
        rec['SERIAL_NO'] = serial_normalize(rec.get('SERIAL_NO'))
        if rec.get('SERIAL_NO') in (None, '', 0, '0', '0.0'):
            history_bad_stone += 1
            continue
        history_clean.append(rec)
    history = history_clean
    if history_bad_stone:
        add_issue(audit, dataset, 'WARNING', 'Blank/zero stone rows in previous final removed', history_path.name,
                  details=f'{history_bad_stone} historical row(s) had blank/0 SERIAL_NO and were excluded.')
    if history_invalid:
        summary = make_failed_summary(dataset, out_name, raw_path, history_path, date_col, f'{history_invalid} invalid history date row(s)')
        audit['summary'].append(summary)
        return summary, False

    history_max = max((r[date_col] for r in history), default=None)
    raw_rows = raw_invalid = blacklist_removed = blank_zero = 0
    multi_stone_rows = expanded_stone_records = 0
    valid_raw = []
    invalid_samples = 0
    for source_row, row in rows_as_dicts(raw_path):
        raw_rows += 1
        rec, _ = map_row(row, dataset, headers, source_map)
        appointment_dt = parse_datetime(rec.get(date_col))
        if appointment_dt is None:
            raw_invalid += 1
            if invalid_samples < 25:
                add_issue(audit, dataset, 'ERROR', 'Invalid raw appointment date', source_label(raw_path), source_row,
                          column=date_col, raw_value=rec.get(date_col),
                          details='Raw View row excluded because APPOINTMENT_DATE cannot be parsed.')
                invalid_samples += 1
            continue
        rec[date_col] = appointment_dt
        if party_rejected(rec):
            blacklist_removed += 1
            continue

        serials = _split_view_serials(rec.get('SERIAL_NO'))
        if not serials:
            blank_zero += 1
            continue
        if len(serials) > 1:
            multi_stone_rows += 1
            expanded_stone_records += len(serials)
        for serial in serials:
            one = dict(rec)
            one['SERIAL_NO'] = serial
            one['_source_row'] = source_row
            valid_raw.append(one)

    carried_history = history
    appended = []
    raw_old = 0
    boundary_history_removed = 0
    if history_max is None:
        appended = valid_raw
    else:
        boundary_date = history_max.date()
        has_boundary_raw = any(r[date_col].date() == boundary_date for r in valid_raw)
        if has_boundary_raw:
            boundary_history_removed = sum(1 for r in history if r[date_col].date() == boundary_date)
            carried_history = [r for r in history if r[date_col].date() < boundary_date]
            appended = [r for r in valid_raw if r[date_col].date() >= boundary_date]
            raw_old = len(valid_raw) - len(appended)
            add_issue(audit, dataset, 'INFO', 'Boundary appointment date refreshed from RAW', source_label(raw_path),
                      details=f'Previous-final rows with APPOINTMENT_DATE {boundary_date} were replaced by current RAW View rows for that date; comma-separated STONE_ID values were split to individual stones.')
        else:
            appended = [r for r in valid_raw if r[date_col] > history_max]
            raw_old = len(valid_raw) - len(appended)

    today = processing_date()
    anchor = today
    cutoff = anchor - timedelta(days=ACTIVITY_DAYS)
    combined = carried_history + appended
    before = after = 0
    filtered = []
    for rec in combined:
        d = rec[date_col].date()
        if d < cutoff:
            before += 1
        elif d > anchor:
            after += 1
        else:
            filtered.append(rec)

    def view_sort_key(r):
        created = parse_datetime(r.get('CREATED_DATETIME')) or datetime.min
        return (r[date_col], created, r.get('_source_row', 0), str(r.get('SERIAL_NO') or ''))
    filtered.sort(key=view_sort_key)

    filtered, exact_duplicates_removed = remove_exact_output_duplicates(filtered, headers)
    if exact_duplicates_removed:
        add_issue(audit, dataset, 'INFO', 'Exact duplicate rows removed',
                  details=f'{exact_duplicates_removed} fully identical View row(s) were removed; the first occurrence was retained.')

    keys = Counter((str(r.get('SEQ_NO')), str(r.get('SERIAL_NO')), str(r.get(date_col))) for r in filtered)
    dups = sum(c - 1 for c in keys.values() if c > 1)
    if dups:
        add_issue(audit, dataset, 'INFO', 'Repeated View activity retained',
                  details=f'{dups} repeated SEQ_NO + SERIAL_NO + APPOINTMENT_DATE occurrence(s) remain because another output field differs; these are not exact duplicate rows.')

    write_standard_output(out_name, headers, filtered)
    audit_blanks(audit, dataset, headers, filtered)
    status = 'PASS' if raw_invalid == 0 and not missing_mapping_rows(mapping_rows) else 'FAILED'
    summary = {
        'Dataset': dataset, 'Output File': out_name, 'Raw File': source_label(raw_path), 'Previous Final File': history_path.name,
        'Source Rule': 'Previous final preserved before latest APPOINTMENT_DATE; boundary appointment date refreshed from RAW when present; later RAW rows appended; comma/semicolon-separated STONE_ID values split into individual SERIAL_NO rows',
        'Window Rule': '90-day window from processing date using APPOINTMENT_DATE; blank/0 SERIAL_NO removed; exact full-row duplicates removed',
        'Date Field': date_col, 'Processing Date': today.isoformat(), 'Anchor Date': anchor.isoformat(), 'Cutoff Date': cutoff.isoformat(),
        'History Rows': len(history), 'History Invalid Dates': history_invalid, 'History Blacklist Removed': history_blacklist, 'Latest History Timestamp': history_max,
        'Raw Rows Read': raw_rows, 'Raw Valid Dates': len(valid_raw), 'Raw Invalid Dates': raw_invalid,
        'Raw New Rows Appended': len(appended), 'Raw Historical/Overlap Rows Skipped': raw_old,
        'Boundary History Rows Replaced': boundary_history_removed,
        'Blacklist Removed From Raw': blacklist_removed, 'Blank/Zero Stone Removed': blank_zero,
        'Multi-Stone Raw Rows Split': multi_stone_rows, 'Expanded Stone Records From Multi-Stone Rows': expanded_stone_records,
        'Before Cutoff': before, 'After Anchor': after, 'Output Rows': len(filtered), 'Exact Duplicate Rows Removed': exact_duplicates_removed, 'Potential Duplicates Kept': dups, 'Status': status
    }
    audit['summary'].append(summary)
    return summary, status == 'PASS'

def carat_group_value(v):
    """Python equivalent of the validated Excel Carat_Grp rule."""
    if is_blank(v): return ''
    try: c = float(v)
    except Exception: return ''
    if c >= 15: return '15+'
    if 8 <= c <= 14.99: return '8.00-14.99'
    if 5 <= c <= 7.99: return '5.00-7.99'
    if 4 <= c <= 4.99: return '4.00-4.99'
    if 3 <= c <= 3.99: return '3.00-3.99'
    if 2.5 <= c <= 2.99: return '2.50-2.99'
    if 2 <= c <= 2.49: return '2.00-2.49'
    if 1.7 <= c <= 1.99: return '1.70-1.99'
    if 1.5 <= c <= 1.69: return '1.50-1.69'
    if 1.2 <= c <= 1.49: return '1.20-1.49'
    if 1 <= c <= 1.19: return '1.00-1.19'
    if 0.9 <= c <= 0.99: return '0.90-0.99'
    if 0.7 <= c <= 0.89: return '0.70-0.89'
    if 0.5 <= c <= 0.69: return '0.50-0.69'
    if 0.4 <= c <= 0.49: return '0.40-0.49'
    if 0.03 <= c <= 0.39: return '0.30-0.39'
    return ''


def sales_amt_value(srate, carat):
    if is_blank(srate) or is_blank(carat): return ''
    try: return float(srate) * float(carat)
    except Exception: return ''


def stock_tag_value(rec: dict):
    parts = [rec.get('Shp'), rec.get('Carat_Grp'), rec.get('Color name'), rec.get('Qua'), rec.get('Lab'), rec.get('Flo')]
    return '-'.join('' if v is None else clean_text(str(v)) for v in parts)


def enrich_sales_derived(rec: dict):
    rec['Carat_Grp'] = carat_group_value(rec.get('Carat'))
    rec['Amt'] = sales_amt_value(rec.get('SRate'), rec.get('Carat'))
    return rec


def enrich_stock_derived(rec: dict):
    rec['Carat_Grp'] = carat_group_value(rec.get('Carat'))
    rec['Tag'] = stock_tag_value(rec)
    return rec


def carat_group_formula(rownum: int, carat_col_letter: str):
    c = f'{carat_col_letter}{rownum}'
    return (f'=IF({c}="","",IF({c}>=15,"15+",IF(AND({c}>=8,{c}<=14.99),"8.00-14.99",'
            f'IF(AND({c}>=5,{c}<=7.99),"5.00-7.99",IF(AND({c}>=4,{c}<=4.99),"4.00-4.99",'
            f'IF(AND({c}>=3,{c}<=3.99),"3.00-3.99",IF(AND({c}>=2.5,{c}<=2.99),"2.50-2.99",'
            f'IF(AND({c}>=2,{c}<=2.49),"2.00-2.49",IF(AND({c}>=1.7,{c}<=1.99),"1.70-1.99",'
            f'IF(AND({c}>=1.5,{c}<=1.69),"1.50-1.69",IF(AND({c}>=1.2,{c}<=1.49),"1.20-1.49",'
            f'IF(AND({c}>=1,{c}<=1.19),"1.00-1.19",IF(AND({c}>=0.9,{c}<=0.99),"0.90-0.99",'
            f'IF(AND({c}>=0.7,{c}<=0.89),"0.70-0.89",IF(AND({c}>=0.5,{c}<=0.69),"0.50-0.69",'
            f'IF(AND({c}>=0.4,{c}<=0.49),"0.40-0.49",IF(AND({c}>=0.03,{c}<=0.39),"0.30-0.39",""))))))))))))))))))')


def sale_key(rec: dict):
    serial = serial_normalize(rec.get('Serial No'))
    dt = parse_datetime(rec.get('Sale Date'))
    return (str(serial), dt) if dt is not None else (str(serial), None)


def read_sales_history(dataset: str, history_path: Path, headers: list[str], audit):
    records = []
    invalid = 0
    blacklist_removed = 0
    invalid_samples = 0
    # Previous Sales 120 may store Sale Date as =INT(P2). Use Excel's cached
    # date value; never parse the formula text as a timestamp.
    for source_row, row in rows_as_dicts(history_path, data_only=True):
        rec = {h: clean_value(row.get(h)) for h in headers if h not in DERIVED_COLUMNS}
        if all(is_blank(v) for v in rec.values()):
            add_issue(audit, dataset, 'INFO', 'Helper-only empty Sales row ignored', history_path.name, source_row)
            continue
        rec['Serial No'] = serial_normalize(rec.get('Serial No'))
        rec['Carat'] = number_normalize(rec.get('Carat'))
        rec['SPer'] = number_normalize(rec.get('SPer'))
        rec['SRate'] = number_normalize(rec.get('SRate'))
        normalize_lab_flo(rec)
        dt = parse_datetime(rec.get('Sale Date'))
        if dt is None:
            invalid += 1
            if invalid_samples < 50:
                add_issue(audit, dataset, 'ERROR', 'Invalid date in previous Sales final', history_path.name, source_row,
                          key=f"Serial No={rec.get('Serial No')}", column='Sale Date', raw_value=rec.get('Sale Date'),
                          details='Sales history row cannot be safely carried forward until its Sale Date is corrected.')
                invalid_samples += 1
            continue
        rec['Sale Date'] = dt
        enrich_sales_derived(rec)
        if party_rejected(rec):
            blacklist_removed += 1
            continue
        rec['_source_row'] = source_row
        records.append(rec)
    return records, invalid, blacklist_removed


def process_sales(raw_path: Path, final120: Path, final180: Path, audit=None):
    dataset = 'Sales'
    # Both Sales outputs always use one canonical schema, regardless of RAW or previous-final header order/names.
    h120 = list(STANDARD_SALES_HEADERS)
    h180 = list(STANDARD_SALES_HEADERS)
    union_headers = list(STANDARD_SALES_HEADERS)
    source_map, mapping_rows = resolve_sales_mapping(raw_path, union_headers)
    audit['mapping'].extend(mapping_rows)

    missing = missing_mapping_rows(mapping_rows)
    record_mapping_issues(audit, dataset, mapping_rows, raw_path)

    if missing:
        reports = []
        for days, out_name in SALES_WINDOWS:
            history_path = final120 if days == 120 else final180
            summary = make_failed_summary(f'Sales {days}', out_name, raw_path, history_path, 'Sale Date',
                                          'Missing required Sales source column(s): ' + ', '.join(m['Final Column'] for m in missing))
            summary['Source Rule'] = 'Processing stopped before output because required raw Sales mappings are missing.'
            audit['summary'].append(summary)
            reports.append(summary)
        return reports, False

    source_headers = [h for h in union_headers if h not in DERIVED_COLUMNS]
    raw_records = []
    raw_dates = []
    raw_rows = invalid = blacklist_removed = mapping_losses = 0
    invalid_samples = 0
    raw_min = raw_max = None
    for source_row, row in rows_as_dicts(raw_path):
        raw_rows += 1
        rec, used = map_row_with_sources(row, source_map, source_headers)
        rec['Serial No'] = serial_normalize(rec.get('Serial No'))
        rec['Carat'] = number_normalize(rec.get('Carat'))
        rec['SPer'] = number_normalize(rec.get('SPer'))
        rec['SRate'] = number_normalize(rec.get('SRate'))
        normalize_lab_flo(rec)
        dt = parse_datetime(rec.get('Sale Date'))
        if dt is None:
            invalid += 1
            if invalid_samples < 25:
                add_issue(audit, dataset, 'ERROR', 'Invalid raw Sale Date', source_label(raw_path), source_row,
                          key=f"Serial No={rec.get('Serial No')}", column='Sale Date', raw_value=rec.get('Sale Date'),
                          details='Raw Sales row cannot be safely processed.')
                invalid_samples += 1
            continue
        rec['Sale Date'] = dt
        enrich_sales_derived(rec)
        raw_dates.append(dt)
        raw_min = dt if raw_min is None or dt < raw_min else raw_min
        raw_max = dt if raw_max is None or dt > raw_max else raw_max
        if party_rejected(rec):
            blacklist_removed += 1
            continue

        # Strong mapping safety: once a RAW source column is resolved, a nonblank source value must survive mapping.
        for target in source_headers:
            source_col = source_map.get(target)
            if not source_col or not is_blank(rec.get(target)):
                continue
            raw_value = row.get(source_col)
            if not is_blank(raw_value):
                mapping_losses += 1
                add_issue(audit, dataset, 'ERROR', 'Sales mapping value lost', source_label(raw_path), source_row,
                          key=key_text(rec, ['Serial No', 'Sale Date']), column=target,
                          raw_value=raw_value,
                          details=f'RAW column {source_col!r} contains a nonblank value but canonical field {target!r} is blank. Output is blocked.')
        rec['_source_row'] = source_row
        raw_records.append(rec)

    today = processing_date()
    anchor = activity_anchor(raw_dates, today)
    if anchor != today:
        add_issue(audit, dataset, 'INFO', 'No current-day Sales log', source_label(raw_path),
                  details=f'No raw Sales record dated {today}; anchor moved to {anchor}.')

    raw_key_counts = Counter(sale_key(r) for r in raw_records)
    raw_duplicate_occurrences = sum(c - 1 for c in raw_key_counts.values() if c > 1)
    if raw_duplicate_occurrences:
        add_issue(audit, dataset, 'WARNING', 'Repeated Sale key in raw kept', source_label(raw_path),
                  details=f'{raw_duplicate_occurrences} repeated Serial No + Sale Date occurrence(s) exist in raw. Final exact full-row duplicates are removed after merge; non-identical repeated sale keys are retained.')

    # Invalid raw dates or lost mapped values are fatal because silently excluding them can corrupt the final window.
    if invalid or mapping_losses:
        reports = []
        reason_bits = []
        if invalid: reason_bits.append(f'{invalid} invalid raw Sale Date row(s)')
        if mapping_losses: reason_bits.append(f'{mapping_losses} raw-to-final mapping value loss(es)')
        for days, out_name in SALES_WINDOWS:
            history_path = final120 if days == 120 else final180
            summary = make_failed_summary(f'Sales {days}', out_name, raw_path, history_path, 'Sale Date', '; '.join(reason_bits))
            summary['Source Rule'] = 'Processing stopped; Sales output is not written when raw rows cannot be mapped safely.'
            audit['summary'].append(summary)
            reports.append(summary)
        return reports, False

    reports = []
    all_ok = True
    for days, out_name in SALES_WINDOWS:
        headers = h120 if days == 120 else h180
        history_path = final120 if days == 120 else final180
        sales_dataset = f'Sales {days}'
        history, history_invalid, history_blacklist = read_sales_history(sales_dataset, history_path, headers, audit)
        cutoff = anchor - timedelta(days=days)

        if history_invalid:
            summary = make_failed_summary(sales_dataset, out_name, raw_path, history_path, 'Sale Date',
                                          f'{history_invalid} invalid previous Sales date row(s)')
            summary['Source Rule'] = 'Processing stopped because previous Sales history contains invalid Sale Date values.'
            audit['summary'].append(summary)
            reports.append(summary)
            all_ok = False
            continue

        history_min = min((r['Sale Date'] for r in history), default=None)
        history_max = max((r['Sale Date'] for r in history), default=None)

        hist_before = hist_after = 0
        history_window = []
        for rec in history:
            d = rec['Sale Date'].date()
            if d < cutoff:
                hist_before += 1
            elif d > anchor:
                hist_after += 1
            else:
                history_window.append(rec)

        raw_before = raw_after = 0
        raw_window = []
        for rec in raw_records:
            d = rec['Sale Date'].date()
            if d < cutoff:
                raw_before += 1
            elif d > anchor:
                raw_after += 1
            else:
                raw_window.append(rec)

        # Sales boundary-date rule: the latest calendar date in the previous Sales base
        # is provisional. When current RAW contains that same date, replace that entire
        # date from RAW and append later RAW dates. This matches the established Jupyter
        # preparation behavior while still removing exact full-row duplicates afterward.
        boundary_history_removed = 0
        raw_overlap_rows = 0
        if history_max is None:
            history_preserved = []
            raw_new_records = list(raw_window)
        else:
            boundary_date = history_max.date()
            has_boundary_raw = any(r['Sale Date'].date() == boundary_date for r in raw_window)
            if has_boundary_raw:
                boundary_history_removed = sum(1 for r in history_window if r['Sale Date'].date() == boundary_date)
                history_preserved = [r for r in history_window if r['Sale Date'].date() < boundary_date]
                raw_new_records = [r for r in raw_window if r['Sale Date'].date() >= boundary_date]
                raw_overlap_rows = sum(1 for r in raw_window if r['Sale Date'].date() == boundary_date)
                add_issue(audit, sales_dataset, 'INFO', 'Boundary Sale Date refreshed from RAW', source_label(raw_path),
                          details=f'Previous Sales rows dated {boundary_date} were removed and that whole date was rebuilt from current RAW; later RAW sales were appended.')
            else:
                history_preserved = list(history_window)
                raw_new_records = [r for r in raw_window if r['Sale Date'] > history_max]
                raw_overlap_rows = len(raw_window) - len(raw_new_records)

        # Kept for backward-compatible reporting fields. A whole-date refresh replaces
        # history rather than mutating individual history rows in-place.
        history_refreshed = []
        raw_new_rows = len(raw_new_records)
        final_records = history_preserved + raw_new_records
        def sk(r):
            try: carat = float(r.get('Carat') or -1)
            except Exception: carat = -1
            return (-carat, r['Sale Date'], r.get('_source_row', 0))
        final_records.sort(key=sk)
        final_records, exact_duplicates_removed = remove_exact_output_duplicates(final_records, headers)
        if exact_duplicates_removed:
            add_issue(audit, sales_dataset, 'INFO', 'Exact duplicate rows removed',
                      details=f'{exact_duplicates_removed} fully identical Sales row(s) were removed; the first occurrence was retained.')

        final_key_counts = Counter(sale_key(r) for r in final_records)
        final_duplicate_occurrences = sum(c - 1 for c in final_key_counts.values() if c > 1)
        if final_duplicate_occurrences:
            add_issue(audit, sales_dataset, 'WARNING', 'Repeated Sale key kept in final',
                      details=f'{final_duplicate_occurrences} repeated Serial No + Sale Date occurrence(s) remain because their complete output rows differ; they are retained.')

        # Coverage safeguard: if neither history nor raw reaches the cutoff, the requested window may be incomplete.
        earliest_available = min([x for x in [history_min, raw_min] if x is not None], default=None)
        if earliest_available is not None and earliest_available.date() > cutoff:
            add_issue(audit, sales_dataset, 'WARNING', 'Sales history may not cover full requested window',
                      source_file=history_path.name,
                      details=f'Earliest available Sale Date is {earliest_available.date()}, but cutoff is {cutoff}.')

        write_sales_output(out_name, headers, final_records)
        audit_blanks(audit, sales_dataset, headers, final_records)

        summary = {
            'Dataset': sales_dataset,
            'Output File': out_name,
            'Raw File': source_label(raw_path),
            'Previous Final File': history_path.name,
            'Source Rule': 'Previous Sales rows before the latest-history calendar date are preserved. If current RAW contains that boundary date, the whole boundary date is replaced from RAW and later RAW dates are appended; otherwise only RAW rows newer than the latest previous timestamp are appended. Exact full-row duplicates are removed.',
            'Window Rule': f'{days}-day window; anchor=today when current-day log exists, otherwise yesterday',
            'Date Field': 'Sale Date',
            'Processing Date': today.isoformat(),
            'Anchor Date': anchor.isoformat(),
            'Cutoff Date': cutoff.isoformat(),
            'History Rows': len(history),
            'History Invalid Dates': history_invalid,
            'Latest History Timestamp': history_max,
            'History Blacklist Removed': history_blacklist,
            'History Rows In Window': len(history_window),
            'History Rows Refreshed By Raw': boundary_history_removed,
            'History Rows Preserved': len(history_preserved),
            'Boundary History Rows Replaced': boundary_history_removed,
            'Raw Rows Read': raw_rows,
            'Raw Valid Dates': len(raw_records),
            'Raw Invalid Dates': invalid,
            'Raw Rows In Window': len(raw_window),
            'Raw Overlap Rows Refreshing History': raw_overlap_rows,
            'Raw New Rows Appended': raw_new_rows,
            'Raw Historical/Overlap Rows Skipped': 0,
            'Blacklist Removed From Raw': blacklist_removed,
            'Before Cutoff': hist_before + raw_before,
            'After Anchor': hist_after + raw_after,
            'Output Rows': len(final_records),
            'Exact Duplicate Rows Removed': exact_duplicates_removed,
            'Potential Duplicates Kept': final_duplicate_occurrences,
            'Status': 'PASS'
        }
        audit['summary'].append(summary)
        reports.append(summary)
    return reports, all_ok and all(r.get('Status') == 'PASS' for r in reports)

def process_stock(raw_path: Path, out_name='Stock file.xlsx', audit=None):
    dataset = 'Stock'
    # Stock is a current snapshot. It does not use or require a previous Stock workbook.
    headers = list(STOCK_FINAL_HEADERS)
    raw_headers = read_headers(raw_path)
    source_map, _ = resolve_standard_mapping(dataset, headers, raw_headers)
    mapping_rows = schema_mapping(dataset, headers, raw_headers)
    audit['mapping'].extend(mapping_rows)
    record_mapping_issues(audit, dataset, mapping_rows, raw_path)
    # Stock Expected is frequently an Excel formula. If that formula is copied into a
    # newly written workbook, openpyxl cannot carry Excel's cached result and downstream
    # data_only readers see a blank. Read the source workbook a second time with
    # data_only=True so the original cached numeric Expected value is preserved.
    cached_rows = {source_row: row for source_row, row in rows_as_dicts(raw_path, data_only=True)}

    records = []
    raw_rows = 0
    nondata_rows_removed = 0
    expected_invalid_rows = 0
    for source_row, row in rows_as_dicts(raw_path):
        raw_rows += 1
        rec, _ = map_row(row, dataset, headers, source_map)
        cached_rec, _ = map_row(cached_rows.get(source_row, {}), dataset, headers, source_map)
        rec['Serial No'] = serial_normalize(rec.get('Serial No'))
        # Some Stock exports append a footer/count row with only Serial No populated.
        # Treat that as non-data rather than a stone record.
        source_fields = [h for h in headers if h not in DERIVED_COLUMNS and h != 'Serial No']
        if not is_blank(rec.get('Serial No')) and all(is_blank(rec.get(h)) for h in source_fields):
            nondata_rows_removed += 1
            add_issue(audit, dataset, 'INFO', 'Serial-only Stock footer/non-data row removed',
                      source_label(raw_path), source_row, key=f"Serial No={rec.get('Serial No')}",
                      details='Only Serial No is populated and all other mapped Stock fields are blank; row is treated as export footer/non-data.')
            continue
        normalize_lab_flo(rec)
        rec['Carat'] = number_normalize(rec.get('Carat'))
        rec['GRap'] = number_normalize(rec.get('GRap'))
        rec['GPer'] = number_normalize(rec.get('GPer'))

        # Prefer the source workbook's cached numeric Expected result when the visible
        # cell is a formula. This keeps the processed base fully value-based, which is
        # required because the Auto Pricing builder reads with data_only=True.
        expected_raw = rec.get('Expected')
        expected_cached = cached_rec.get('Expected')
        if isinstance(expected_raw, str) and expected_raw.lstrip().startswith('='):
            rec['Expected'] = number_normalize(expected_cached)
        else:
            rec['Expected'] = number_normalize(expected_raw)

        # Defensive fallback for formula workbooks whose cached values were stripped by
        # an upstream export. Current Stock business formula is:
        # GRap * (100 - (GPer + 5))% * (100 - 10)%.
        if is_blank(rec.get('Expected')) and not is_blank(rec.get('GRap')) and not is_blank(rec.get('GPer')):
            try:
                rec['Expected'] = float(rec['GRap']) * ((100.0 - (float(rec['GPer']) + 5.0)) / 100.0) * 0.90
                add_issue(audit, dataset, 'WARNING', 'Expected cache missing; value recalculated', source_label(raw_path), source_row,
                          key=f"Serial No={rec.get('Serial No')}", column='Expected', raw_value=expected_raw,
                          details='The Excel formula had no cached result, so Expected was recalculated from GRap/GPer using the current Stock business formula.')
            except Exception:
                pass
        if (not isinstance(rec.get('Expected'), (int, float)) or isinstance(rec.get('Expected'), bool)
                or not math.isfinite(float(rec['Expected']))):
            expected_invalid_rows += 1
            if expected_invalid_rows <= 25:
                add_issue(audit, dataset, 'ERROR', 'Expected cannot be calculated safely', source_label(raw_path), source_row,
                          key=f"Serial No={rec.get('Serial No')}", column='Expected',
                          details='Provide numeric Expected, or both numeric GRap and GPer. Stock is not marked PASS when Expected is unavailable.')
        enrich_stock_derived(rec)
        rec['_source_row'] = source_row
        records.append(rec)
    def sk(r):
        try: carat = float(r.get('Carat') or -1)
        except Exception: carat = -1
        return (-carat, r.get('_source_row', 0))
    records.sort(key=sk)
    records, exact_duplicates_removed = remove_exact_output_duplicates(records, headers)
    if exact_duplicates_removed:
        add_issue(audit, dataset, 'INFO', 'Exact duplicate rows removed', source_label(raw_path),
                  details=f'{exact_duplicates_removed} fully identical Stock row(s) were removed; the first occurrence was retained.')
    serial_counts = Counter(str(r.get('Serial No')) for r in records if not is_blank(r.get('Serial No')))
    dups = sum(c - 1 for c in serial_counts.values() if c > 1)
    if dups:
        add_issue(audit, dataset, 'WARNING', 'Potential duplicate Serial No kept', source_label(raw_path),
                  details=f'{dups} repeated Serial No occurrence(s) found. No deduplication is applied.')
    write_stock_output(out_name, headers, records)
    audit_blanks(audit, dataset, headers, records)
    status = ('PASS' if expected_invalid_rows == 0 and not missing_mapping_rows(mapping_rows)
              else 'FAILED')
    summary = {
        'Dataset': dataset, 'Output File': out_name, 'Raw File': source_label(raw_path),
        'Previous Final File': '', 'Source Rule': 'RAW Stock only; no previous Stock file required; formula-backed Expected values are materialized from source cached values; no history merge.',
        'Window Rule': 'Current snapshot', 'Date Field': '', 'Processing Date': processing_date().isoformat(), 'Anchor Date': '', 'Cutoff Date': '',
        'History Rows': 'Not used', 'History Invalid Dates': 'Not used', 'Latest History Timestamp': 'Not used',
        'Raw Rows Read': raw_rows, 'Raw Valid Dates': '', 'Raw Invalid Dates': 0, 'Raw New Rows Appended': 'All valid raw rows are source',
        'Raw Historical/Overlap Rows Skipped': 0, 'Blacklist Removed From Raw': 0, 'Stock Non-Data Rows Removed': nondata_rows_removed,
        'Before Cutoff': 0, 'After Anchor': 0,
        'Output Rows': len(records), 'Exact Duplicate Rows Removed': exact_duplicates_removed, 'Potential Duplicates Kept': dups,
        'Stock Expected Invalid Rows': expected_invalid_rows, 'Status': status
    }
    audit['summary'].append(summary)
    return summary, status == 'PASS'


def make_failed_summary(dataset, out_name, raw_path, history_path, date_col, reason):
    return {
        'Dataset': dataset, 'Output File': out_name, 'Raw File': source_label(raw_path), 'Previous Final File': history_path.name,
        'Source Rule': 'Processing stopped for this dataset', 'Window Rule': '', 'Date Field': date_col,
        'Processing Date': processing_date().isoformat(), 'Anchor Date': '', 'Cutoff Date': '',
        'History Rows': '', 'History Invalid Dates': reason, 'Latest History Timestamp': '', 'Raw Rows Read': '',
        'Raw Valid Dates': '', 'Raw Invalid Dates': '', 'Raw New Rows Appended': '', 'Raw Historical/Overlap Rows Skipped': '',
        'Blacklist Removed From Raw': '', 'Before Cutoff': '', 'After Anchor': '', 'Output Rows': '',
        'Potential Duplicates Kept': '', 'Status': 'FAILED'
    }


def audit_blanks(audit, dataset: str, headers: list[str], records: list[dict]):
    total = len(records)
    for h in headers:
        if h in DERIVED_COLUMNS: continue
        blanks = sum(1 for r in records if is_blank(r.get(h)))
        audit['blanks'].append({'Dataset': dataset, 'Final Column': h, 'Output Rows': total, 'Blank Count': blanks,
                                'Nonblank Count': total - blanks, 'Blank %': blanks / total if total else 0,
                                'Comment': 'Blank remains blank only when source/history value is blank after CLEAN/TRIM.'})


def write_standard_output(out_name: str, headers: list[str], records: list[dict]):
    wb = Workbook(write_only=True); ws = wb.create_sheet('Sheet1'); ws.append(headers)
    for rec in records: ws.append([rec.get(h) for h in headers])
    wb.save(OUT_DIR / out_name)


def write_sales_output(out_name: str, headers: list[str], records: list[dict]):
    # v4.1.1: derived fields are already calculated in Python. Write values directly
    # so downstream application logic never depends on Excel cached formula results.
    wb = Workbook(write_only=True); ws = wb.create_sheet('Sheet1'); ws.append(headers)
    for rec in records:
        ws.append([rec.get(h) for h in headers])
    wb.save(OUT_DIR / out_name)


def write_stock_output(out_name: str, headers: list[str], records: list[dict]):
    # v4.1.1: Carat_Grp and Tag are calculated in Python and persisted as values.
    wb = Workbook(write_only=True); ws = wb.create_sheet('Sheet1'); ws.append(headers)
    for rec in records:
        ws.append([rec.get(h) for h in headers])
    wb.save(OUT_DIR / out_name)

def write_audit(audit, path: Path):
    wb = Workbook(); wb.remove(wb.active)
    for sheet_name, rows in [('Summary', audit['summary']), ('Column Mapping', audit['mapping']), ('Blank Analysis', audit['blanks']), ('Issues', audit['issues'])]:
        ws = wb.create_sheet(sheet_name)
        if not rows:
            ws.append(['No records']); continue
        headers = list(dict.fromkeys(k for r in rows for k in r.keys())); ws.append(headers)
        for r in rows: ws.append([r.get(h) for h in headers])
        ws.freeze_panes = 'A2'; ws.auto_filter.ref = ws.dimensions
        for c in ws[1]:
            c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='1F4E78'); c.alignment = Alignment(vertical='center')
        for idx, h in enumerate(headers, 1):
            width = min(55, max(12, len(str(h)) + 2))
            for row in range(2, min(ws.max_row, 150) + 1):
                v = ws.cell(row, idx).value
                if v is not None: width = min(55, max(width, min(55, len(str(v)) + 2)))
            ws.column_dimensions[get_column_letter(idx)].width = width
        if sheet_name == 'Blank Analysis' and 'Blank %' in headers:
            ci = headers.index('Blank %') + 1
            for r in range(2, ws.max_row + 1): ws.cell(r, ci).number_format = '0.00%'
        if sheet_name == 'Issues' and 'Severity' in headers:
            ci = headers.index('Severity') + 1
            fills = {'ERROR':'F4CCCC','WARNING':'FCE5CD','INFO':'D9EAD3'}
            for r in range(2, ws.max_row + 1):
                sev = ws.cell(r, ci).value
                if sev in fills: ws.cell(r, ci).fill = PatternFill('solid', fgColor=fills[sev])
    wb.save(path)


def main():
    # The pipeline sets OUTPUT_DIR to a writable per-run temporary directory.
    # Only an actual processing run needs this directory, never an import.
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    audit = {'summary': [], 'mapping': [], 'blanks': [], 'issues': []}
    ok_flags = []
    cart_raw = locate([r'ADD[_ -]?TO[_ -]?CART', r'Cart'], RAW_DIR)
    hold_raw = locate([r'JANGAD[_ -]?HISTORY', r'^JANGAD$', r'Hold'], RAW_DIR)
    offer_raw = locate([r'OFFER[_ -]?PLACE', r'Offer'], RAW_DIR)
    view_raw = locate([r'APPOINTMENT', r'View'], RAW_DIR)
    wishlist_raw = locate([r'WISHLIST', r'Wish'], RAW_DIR)
    sales_raw = locate_sales(RAW_DIR)
    stock_raw = locate([r'^Stock$', r'^KSTAKA', r'^KSSTAKA', r'^KSATAKA'], RAW_DIR)

    _, ok = process_history_log('Cart', cart_raw, final_path('Cart.xlsx'), 'Cart.xlsx', 'CREATED_DATETIME', ACTIVITY_DAYS, audit=audit); ok_flags.append(ok)
    _, ok = process_history_log('Hold', hold_raw, final_path('Hold.xlsx'), 'Hold.xlsx', 'JANGAD_DATE', HOLD_DAYS, hold_mode=True, audit=audit); ok_flags.append(ok)
    _, ok = process_history_log('Offer', offer_raw, final_path('Offer.xlsx'), 'Offer.xlsx', 'CREATED_DATETIME', ACTIVITY_DAYS, audit=audit); ok_flags.append(ok)
    _, ok = process_view(view_raw, final_path('View.xlsx'), 'View.xlsx', audit=audit); ok_flags.append(ok)
    _, ok = process_history_log('Wishlist', wishlist_raw, final_path('Wishlist.xlsx'), 'Wishlist.xlsx', 'CREATED_DATETIME', ACTIVITY_DAYS, audit=audit); ok_flags.append(ok)
    _, ok = process_sales(sales_raw, final_path('sales_120_file.xlsx'), final_path('sales_180_file.xlsx'), audit=audit); ok_flags.append(ok)
    _, ok = process_stock(stock_raw, 'Stock file.xlsx', audit=audit); ok_flags.append(ok)

    audit_path = OUT_DIR / 'Processing_Audit.xlsx'; write_audit(audit, audit_path)
    report = {
        'processor_version': VERSION, 'processing_date': processing_date().isoformat(),
        'rules': {
            'Processing Date': 'Production uses actual current date in Asia/Kolkata. TEST_PROCESSING_DATE is accepted only for regression/testing.',
            'Legacy Date Column': 'Obsolete final helper column Date is ignored/removed; it is not required from RAW and is not used for date-window logic',
            'Cart/Offer/Wishlist': 'Previous final preserved before its last calendar date; if RAW contains that date, replace the entire boundary date from RAW; otherwise append after latest timestamp; 90-day window; exact full-row duplicates removed',
            'View': 'Previous final preserved before the latest APPOINTMENT_DATE; boundary appointment date refreshed from RAW; comma/semicolon-separated STONE_ID values are split to individual SERIAL_NO rows; blank/0 removed; 90-day window uses APPOINTMENT_DATE; exact full-row duplicates removed',
            'Hold': 'Same whole-boundary-date refresh rule; SERIAL_NO comes from RAW SERIAL_NO; 90-day window from processing date; exact full-row duplicates removed; PARTY_CODE is optional and never blocks Auto Pricing',
            'Sales': "Sales 120/180 previous files carry descriptive history before the boundary date. When current RAW contains the previous file's latest Sale Date, that whole date is replaced from RAW and later RAW dates append; otherwise only rows newer than the latest previous timestamp append. Blank Lab/Flo become NON.",
            'Stock': 'RAW Stock only; no previous Stock file required; canonical output headers are built into the processor; formula-backed Expected is materialized as a numeric value from the source cache (with a GRap/GPer fallback); exact full-row duplicates are removed',
            'Duplicate Cleaning': 'Only fully identical output rows are removed after merge/window processing. Similar or repeated activity with any differing field is retained.',
            'Derived Fields': 'Sales Carat_Grp/Amt and Stock Carat_Grp/Tag are computed in Python and written as values; blank Lab/Flo become NON (including blank Stock Lab -> NON); downstream logic does not depend on Excel cached formulas.',
            'Missing Optional Columns': 'Missing descriptive fields not used for pricing/record matching/window selection are reported as WARNING and blank in newly imported records. Critical fields still block.',
        },
        'summary': audit['summary'], 'issue_counts': dict(Counter(x['Severity'] for x in audit['issues'])),
        'all_datasets_ok': all(ok_flags), 'audit_file': audit_path.name
    }
    report_path = OUT_DIR / 'processing_report.json'; report_path.write_text(json.dumps(report, indent=2, default=str), encoding='utf-8')

    zname = 'final_processed_files_v4.zip' if all(ok_flags) else 'processing_review_v4.zip'
    zpath = OUT_DIR / zname
    with zipfile.ZipFile(zpath, 'w', zipfile.ZIP_DEFLATED) as z:
        for name in ['Cart.xlsx','Hold.xlsx','Offer.xlsx','View.xlsx','Wishlist.xlsx','sales_120_file.xlsx','sales_180_file.xlsx','Stock file.xlsx','Processing_Audit.xlsx']:
            p = OUT_DIR / name
            if p.exists(): z.write(p, p.name)
        z.write(report_path, report_path.name)
    print(json.dumps(report, indent=2, default=str))
    print('ZIP', zpath)
    if not all(ok_flags):
        print('REVIEW_REQUIRED: One or more datasets have errors. Open Processing_Audit.xlsx before using outputs.')


if __name__ == '__main__': main()
