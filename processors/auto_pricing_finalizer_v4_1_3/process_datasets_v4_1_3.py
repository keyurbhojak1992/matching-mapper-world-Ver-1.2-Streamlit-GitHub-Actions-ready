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

VERSION = '4.1.3'
RAW_DIR = Path(os.environ.get('RAW_DATA_DIR', '/mnt/data/raw_inputs'))
FINAL_DIR = Path(os.environ.get('FINAL_HISTORY_DIR', '/mnt/data/final_samples'))
OUT_DIR = Path(os.environ.get('OUTPUT_DIR', '/mnt/data/output_final'))
OUT_DIR.mkdir(parents=True, exist_ok=True)

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
STOCK_FINAL_HEADERS = ['Serial No', 'Shp', 'Carat', 'Carat_Grp', 'Color name', 'Qua', 'Cut', 'Pol', 'Sym', 'Lab', 'Flo', 'GRap', 'GPer', 'Expected', 'Tag']
LEGACY_FINAL_COLUMNS = {'date'}  # Obsolete helper column; never used for current-date/window logic.
MAPPING_FILE = Path(os.environ.get('COLUMN_MAPPING_FILE', str(Path(__file__).with_name('column_mapping_aliases.json'))))

ALIASES = {
    'Hold': {
        'SERIAL_NO': ['SERIAL_NO', 'Serial No', 'STONE_ID'],
        'JANGAD_DATE': ['JANGAD_DATE', 'Date', 'date'],
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
        'Sale Date': ['Sale Date', 'Sales Date', 'Sold Date', 'Date Sold', 'S date', 'S_DATE', 'SDATE', 'SALE_DATE', 'SALE_DT', 'SALES_DATE', 'SOLD_DATE', 'Invoice Date', 'Bill Date', 'Transaction Date', 'Txn Date'],
        'SPer': ['SPer', 'SALE_PER', 'Sale %', 'Sale Percent', 'Discount %', 'Discount', 'S.%', 'S%', 'SALEPER', 'Sale Per'],
        'SRate': ['SRate', 'Rate', 'Sale Rate', 'Sales Rate', 'SALE_RATE', 'S RATE', 'SRATE', 'PRICE_RATE'],
    },
    'Stock': {
        'Serial No': ['Serial No', 'SERIAL_NO', 'NO.', 'STONE_ID'], 'Shp': ['Shp', 'Shape'],
        'Carat': ['Carat', 'CTS', 'WEIGHT'], 'Color name': ['Color name', 'Color', 'Col'],
        'Qua': ['Qua', 'Quality', 'Clarity'], 'Cut': ['Cut'], 'Pol': ['Pol'], 'Sym': ['Sym'],
        'Lab': ['Lab'], 'Flo': ['Flo', 'FL'], 'GRap': ['GRap', 'RAP'],
        'GPer': ['GPer', 'GPER', 'G.%'], 'Expected': ['Expected', 'Exp rate', 'Expected Rate', 'EXP_RATE'],
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


def rows_as_dicts(path: Path):
    wb = load_workbook(path, read_only=True, data_only=False)
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


def resolve_sales_mapping(raw_path: Path, final_headers: list[str] | None = None):
    """Resolve RAW Sales headers once, then map every row to fixed canonical final names."""
    final_headers = list(final_headers or STANDARD_SALES_HEADERS)
    raw_headers = read_headers(raw_path)
    by_norm = defaultdict(list)
    for h in raw_headers:
        if h:
            by_norm[norm_header(h)].append(h)
    resolved = {}
    used = set()
    rows = []

    # Pass 1: exact alias match (case/space/punctuation insensitive).
    for target in final_headers:
        if target in DERIVED_COLUMNS:
            rows.append({'Dataset': 'Sales', 'Final Column': target, 'Candidate Raw Columns': '', 'Raw Columns Found': '',
                         'Status': 'DERIVED', 'Rule': 'Calculated in output', 'Mapping Method': 'DERIVED', 'Confidence': 1.0})
            continue
        found = []
        for cand in candidates_for('Sales', target):
            found.extend(by_norm.get(norm_header(cand), []))
        found = [h for h in dict.fromkeys(found) if h not in used]
        if found:
            source = found[0]
            resolved[target] = source
            used.add(source)
            rows.append({'Dataset': 'Sales', 'Final Column': target, 'Candidate Raw Columns': ' | '.join(candidates_for('Sales', target)),
                         'Raw Columns Found': source, 'Status': 'OK', 'Rule': 'Canonical alias match; first nonblank source value used',
                         'Mapping Method': 'EXACT_ALIAS', 'Confidence': 1.0})

    # Pass 2: special Sale Date detection from actual values/Excel date formatting.
    if 'Sale Date' not in resolved and 'Sale Date' in final_headers:
        source, confidence, method = detect_sale_date_column(raw_path, raw_headers, used)
        if source:
            resolved['Sale Date'] = source
            used.add(source)
            rows = [r for r in rows if r['Final Column'] != 'Sale Date']
            rows.append({'Dataset': 'Sales', 'Final Column': 'Sale Date', 'Candidate Raw Columns': ' | '.join(candidates_for('Sales', 'Sale Date')),
                         'Raw Columns Found': source, 'Status': 'OK',
                         'Rule': 'Unknown header recognized from date-like cell values and/or Excel date number format',
                         'Mapping Method': method, 'Confidence': round(confidence, 4)})

    # Pass 3: semantic header hints, then conservative fuzzy alias matching.
    for target in final_headers:
        if target in DERIVED_COLUMNS or target in resolved:
            continue
        chosen = None
        method = None
        confidence = 0.0
        hints = SALES_HEADER_HINTS.get(target, ())
        for raw in raw_headers:
            if raw in used:
                continue
            nr = norm_header(raw)
            if nr and any(hint in nr for hint in hints):
                chosen = raw
                method = 'HEADER_HINT'
                confidence = 0.90
                break
        if chosen is None:
            fuzzy = _best_fuzzy_sales_header(target, raw_headers, used)
            if fuzzy:
                confidence, chosen, _alias = fuzzy
                method = 'FUZZY_ALIAS'
        if chosen:
            resolved[target] = chosen
            used.add(chosen)
            rows.append({'Dataset': 'Sales', 'Final Column': target, 'Candidate Raw Columns': ' | '.join(candidates_for('Sales', target)),
                         'Raw Columns Found': chosen, 'Status': 'OK', 'Rule': 'Alternate header normalized to canonical final name',
                         'Mapping Method': method, 'Confidence': round(confidence, 4)})
        else:
            rows.append({'Dataset': 'Sales', 'Final Column': target, 'Candidate Raw Columns': ' | '.join(candidates_for('Sales', target)),
                         'Raw Columns Found': '', 'Status': 'MISSING',
                         'Rule': 'No safe exact/semantic/fuzzy mapping found; processor fails rather than guessing',
                         'Mapping Method': 'MISSING', 'Confidence': 0.0})

    # Preserve canonical final column order in audit.
    order = {h: i for i, h in enumerate(final_headers)}
    rows.sort(key=lambda r: order.get(r['Final Column'], 999))
    return resolved, rows


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


def map_row(row: dict, dataset: str, final_headers: list[str]):
    nmap = defaultdict(list)
    for k, v in row.items():
        nk = norm_header(k)
        if nk: nmap[nk].append((k, v))
    out = {}
    source_used = {}
    for target in final_headers:
        if target in DERIVED_COLUMNS: continue
        chosen = None
        chosen_source = None
        for cand in candidates_for(dataset, target):
            for source_name, value in nmap.get(norm_header(cand), []):
                if not is_blank(value):
                    chosen = clean_value(value)
                    chosen_source = source_name
                    break
            if chosen_source is not None: break
        out[target] = chosen
        source_used[target] = chosen_source
    return out, source_used


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


def schema_mapping(dataset: str, final_headers: list[str], raw_headers: list[str]):
    raw_norm_to_names = defaultdict(list)
    for h in raw_headers: raw_norm_to_names[norm_header(h)].append(h)
    rows = []
    for target in final_headers:
        if target in DERIVED_COLUMNS:
            rows.append({'Dataset': dataset, 'Final Column': target, 'Candidate Raw Columns': '', 'Raw Columns Found': '',
                         'Status': 'DERIVED', 'Rule': 'Calculated in output'})
            continue
        candidates = candidates_for(dataset, target)
        found = []
        for cand in candidates: found.extend(raw_norm_to_names.get(norm_header(cand), []))
        found = list(dict.fromkeys(found))
        rows.append({'Dataset': dataset, 'Final Column': target, 'Candidate Raw Columns': ' | '.join(candidates),
                     'Raw Columns Found': ' | '.join(found), 'Status': 'OK' if found else 'MISSING',
                     'Rule': 'First nonblank candidate used; blank/whitespace candidate is skipped'})
    return rows


def read_history(dataset, history_path: Path, headers, date_col, audit):
    records = []
    invalid = 0
    blacklist_removed = 0
    samples = 0
    blacklist_samples = 0
    for source_row, row in rows_as_dicts(history_path):
        rec = {h: clean_value(row.get(h)) for h in headers}
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


def process_history_log(dataset: str, raw_path: Path, history_path: Path, out_name: str, date_col: str,
                        days: int, hold_mode=False, audit=None):
    """History + RAW processor with authoritative whole-boundary-date refresh.

    If RAW contains the last calendar date present in the previous final, every previous-final
    row from that date is discarded and that whole date is rebuilt from RAW. This protects
    against partial-day previous finals that ended at a timestamp before the day was complete.
    """
    headers = read_final_headers(history_path)
    raw_headers = read_headers(raw_path)
    mapping_rows = schema_mapping(dataset, headers, raw_headers)
    audit['mapping'].extend(mapping_rows)
    for m in mapping_rows:
        if m['Status'] == 'MISSING':
            add_issue(audit, dataset, 'ERROR', 'Missing source column', source_label(raw_path), column=m['Final Column'],
                      details=f"No raw source found. Candidates: {m['Candidate Raw Columns']}")

    history, history_invalid, history_blacklist = read_history(dataset, history_path, headers, date_col, audit)
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
        rec, _ = map_row(row, dataset, headers)
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

    # Audit duplicates only; do not remove them in v4.1.3.
    sig_cols = [h for h in headers if h not in DERIVED_COLUMNS]
    sig = lambda rec: tuple(str(rec.get(h)) for h in sig_cols)
    counts = Counter(sig(rec) for rec in filtered)
    dups = sum(c - 1 for c in counts.values() if c > 1)
    new_dups = ''
    if dataset == 'Hold':
        hist_in_window = [r for r in carried_history if cutoff <= r[date_col].date() <= anchor]
        app_in_window = [r for r in appended if cutoff <= r[date_col].date() <= anchor]
        hc = Counter(sig(r) for r in hist_in_window)
        ac = Counter(sig(r) for r in app_in_window)
        new_dups = sum(max(hc[k] + ac[k] - 1, 0) - max(hc[k] - 1, 0) for k in set(hc) | set(ac))
        if new_dups:
            add_issue(audit, dataset, 'WARNING', 'New duplicate-looking occurrences this run',
                      details=f'{new_dups} duplicate-looking occurrence(s) were introduced by newly appended/refreshed RAW rows. Rows are kept; no deduplication is applied.')
        if dups:
            add_issue(audit, dataset, 'INFO', 'Total duplicate-looking occurrences in final',
                      details=f'{dups} total duplicate-looking occurrence(s) exist in the final Hold window, including historical occurrences. Rows are kept.')
    elif dups:
        add_issue(audit, dataset, 'WARNING', 'Potential duplicate-looking rows kept',
                  details=f'{dups} repeated occurrence(s) kept. Exact row deduplication is deferred.')

    write_standard_output(out_name, headers, filtered)
    audit_blanks(audit, dataset, headers, filtered)
    status = 'PASS' if raw_invalid == 0 and all(m['Status'] != 'MISSING' for m in mapping_rows) else 'FAILED'
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
        'Output Rows': len(filtered), 'Potential Duplicates Kept': dups,
        'New Duplicate-Looking Occurrences This Run': new_dups if dataset == 'Hold' else '',
        'Total Duplicate-Looking Occurrences In Final': dups if dataset == 'Hold' else '', 'Status': status
    }
    audit['summary'].append(summary)
    return summary, status == 'PASS'

def process_view(raw_path: Path, history_path: Path, out_name='View.xlsx', audit=None):
    dataset = 'View'
    headers = read_final_headers(history_path)
    raw_headers = read_headers(raw_path)
    mapping_rows = schema_mapping(dataset, headers, raw_headers)
    audit['mapping'].extend(mapping_rows)
    for m in mapping_rows:
        if m['Status'] == 'MISSING':
            add_issue(audit, dataset, 'ERROR', 'Missing source column', source_label(raw_path), column=m['Final Column'],
                      details=f"No raw source found. Candidates: {m['Candidate Raw Columns']}")

    history, history_invalid, history_blacklist = read_history(dataset, history_path, headers, 'CREATED_DATETIME', audit)
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
        summary = make_failed_summary(dataset, out_name, raw_path, history_path, 'CREATED_DATETIME', f'{history_invalid} invalid history date row(s)')
        audit['summary'].append(summary)
        return summary, False

    history_max = max((r['CREATED_DATETIME'] for r in history), default=None)
    raw_rows = raw_invalid = blacklist_removed = blank_zero = 0
    valid_raw = []
    invalid_samples = 0
    for source_row, row in rows_as_dicts(raw_path):
        raw_rows += 1
        rec, _ = map_row(row, dataset, headers)
        rec['SERIAL_NO'] = serial_normalize(rec.get('SERIAL_NO'))
        if rec.get('SERIAL_NO') in (None, '', 0, '0', '0.0'):
            blank_zero += 1
            continue
        dt = parse_datetime(rec.get('CREATED_DATETIME'))
        if dt is None:
            raw_invalid += 1
            if invalid_samples < 25:
                add_issue(audit, dataset, 'ERROR', 'Invalid raw date', source_label(raw_path), source_row,
                          column='CREATED_DATETIME', raw_value=rec.get('CREATED_DATETIME'),
                          details='Raw View row excluded because date cannot be parsed.')
                invalid_samples += 1
            continue
        rec['CREATED_DATETIME'] = dt
        rec['_source_row'] = source_row
        if party_rejected(rec):
            blacklist_removed += 1
            continue
        valid_raw.append(rec)

    carried_history = history
    appended = []
    raw_old = 0
    boundary_history_removed = 0
    if history_max is None:
        appended = valid_raw
    else:
        boundary_date = history_max.date()
        has_boundary_raw = any(r['CREATED_DATETIME'].date() == boundary_date for r in valid_raw)
        if has_boundary_raw:
            boundary_history_removed = sum(1 for r in history if r['CREATED_DATETIME'].date() == boundary_date)
            carried_history = [r for r in history if r['CREATED_DATETIME'].date() < boundary_date]
            appended = [r for r in valid_raw if r['CREATED_DATETIME'].date() >= boundary_date]
            raw_old = len(valid_raw) - len(appended)
            add_issue(audit, dataset, 'INFO', 'Boundary date refreshed from RAW', source_label(raw_path),
                      details=f'Previous-final rows dated {boundary_date} were replaced by current RAW View rows for that date. STONE_ID is kept as one SERIAL_NO value; it is not split.')
        else:
            appended = [r for r in valid_raw if r['CREATED_DATETIME'] > history_max]
            raw_old = len(valid_raw) - len(appended)

    today = processing_date()
    anchor = today
    cutoff = anchor - timedelta(days=ACTIVITY_DAYS)
    combined = carried_history + appended
    before = after = 0
    filtered = []
    for rec in combined:
        d = rec['CREATED_DATETIME'].date()
        if d < cutoff: before += 1
        elif d > anchor: after += 1
        else: filtered.append(rec)
    filtered.sort(key=lambda r: (r['CREATED_DATETIME'], r.get('_source_row', 0), str(r.get('SERIAL_NO') or '')))

    keys = Counter((str(r.get('SEQ_NO')), str(r.get('SERIAL_NO')), str(r.get('CREATED_DATETIME'))) for r in filtered)
    dups = sum(c - 1 for c in keys.values() if c > 1)
    if dups:
        add_issue(audit, dataset, 'WARNING', 'Potential View duplicates kept', details=f'{dups} repeated SEQ_NO + SERIAL_NO + CREATED_DATETIME occurrence(s) kept. Exact row deduplication is deferred to the later deduplication step.')

    write_standard_output(out_name, headers, filtered)
    audit_blanks(audit, dataset, headers, filtered)
    status = 'PASS' if raw_invalid == 0 and all(m['Status'] != 'MISSING' for m in mapping_rows) else 'FAILED'
    summary = {
        'Dataset': dataset, 'Output File': out_name, 'Raw File': source_label(raw_path), 'Previous Final File': history_path.name,
        'Source Rule': 'Previous final preserved before latest-history date; boundary date refreshed from RAW when present; later RAW rows appended; STONE_ID kept as one SERIAL_NO value',
        'Window Rule': '90-day window from processing date; blank/0 SERIAL_NO removed; no STONE_ID split; exact row deduplication deferred',
        'Date Field': 'CREATED_DATETIME', 'Processing Date': today.isoformat(), 'Anchor Date': anchor.isoformat(), 'Cutoff Date': cutoff.isoformat(),
        'History Rows': len(history), 'History Invalid Dates': history_invalid, 'History Blacklist Removed': history_blacklist, 'Latest History Timestamp': history_max,
        'Raw Rows Read': raw_rows, 'Raw Valid Dates': len(valid_raw), 'Raw Invalid Dates': raw_invalid,
        'Raw New Rows Appended': len(appended), 'Raw Historical/Overlap Rows Skipped': raw_old,
        'Boundary History Rows Replaced': boundary_history_removed,
        'Blacklist Removed From Raw': blacklist_removed, 'Blank/Zero Stone Removed': blank_zero,
        'Before Cutoff': before, 'After Anchor': after, 'Output Rows': len(filtered), 'Potential Duplicates Kept': dups, 'Status': status
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
    for source_row, row in rows_as_dicts(history_path):
        rec = {h: clean_value(row.get(h)) for h in headers if h not in DERIVED_COLUMNS}
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

    missing = [m for m in mapping_rows if m['Status'] == 'MISSING' and m['Final Column'] not in DERIVED_COLUMNS]
    for m in missing:
        add_issue(audit, dataset, 'ERROR', 'Missing required Sales source column', source_label(raw_path), column=m['Final Column'],
                  details=f"No raw source found. Candidates: {m['Candidate Raw Columns']}")

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
                  details=f'{raw_duplicate_occurrences} repeated Serial No + Sale Date occurrence(s) exist in raw. They are kept; no raw deduplication is applied.')

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

        # Sales overlap rule: previous input remains the descriptive base.
        # For an exact Serial No + Sale Date overlap, current RAW updates only SPer/SRate
        # (and therefore Amt). New RAW transactions are appended with their full mapped values.
        history_key_set = {sale_key(r) for r in history_window}
        raw_first = {}
        for r in raw_window:
            raw_first.setdefault(sale_key(r), r)
        history_refreshed = []
        history_preserved = []
        for hrec in history_window:
            rr = raw_first.get(sale_key(hrec))
            if rr is None:
                history_preserved.append(hrec)
                continue
            upd = dict(hrec)
            upd['SPer'] = rr.get('SPer')
            upd['SRate'] = rr.get('SRate')
            normalize_lab_flo(upd)
            enrich_sales_derived(upd)
            history_refreshed.append(upd)
        raw_overlap_rows = sum(1 for r in raw_window if sale_key(r) in history_key_set)
        raw_new_records = [r for r in raw_window if sale_key(r) not in history_key_set]
        raw_new_rows = len(raw_new_records)

        final_records = history_preserved + history_refreshed + raw_new_records
        def sk(r):
            try: carat = float(r.get('Carat') or -1)
            except Exception: carat = -1
            return (-carat, r['Sale Date'], r.get('_source_row', 0))
        final_records.sort(key=sk)

        final_key_counts = Counter(sale_key(r) for r in final_records)
        final_duplicate_occurrences = sum(c - 1 for c in final_key_counts.values() if c > 1)
        if final_duplicate_occurrences:
            add_issue(audit, sales_dataset, 'WARNING', 'Repeated Sale key kept in final',
                      details=f'{final_duplicate_occurrences} repeated Serial No + Sale Date occurrence(s) remain. No deduplication is applied.')

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
            'Source Rule': 'Previous Sales input remains the descriptive base. Current RAW updates SPer/SRate on exact Serial No + Sale Date overlaps; new RAW transactions are appended. Exact row deduplication is deferred.',
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
            'History Rows Refreshed By Raw': len(history_refreshed),
            'History Rows Preserved': len(history_preserved),
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
    mapping_rows = schema_mapping(dataset, headers, raw_headers)
    audit['mapping'].extend(mapping_rows)
    for m in mapping_rows:
        if m['Status'] == 'MISSING' and m['Final Column'] not in DERIVED_COLUMNS:
            add_issue(audit, dataset, 'ERROR', 'Missing source column', source_label(raw_path), column=m['Final Column'],
                      details=f"No raw source found. Candidates: {m['Candidate Raw Columns']}")
    records = []
    raw_rows = 0
    nondata_rows_removed = 0
    for source_row, row in rows_as_dicts(raw_path):
        raw_rows += 1
        rec, _ = map_row(row, dataset, headers)
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
        rec['Expected'] = number_normalize(rec.get('Expected'))
        enrich_stock_derived(rec)
        rec['_source_row'] = source_row
        records.append(rec)
    def sk(r):
        try: carat = float(r.get('Carat') or -1)
        except Exception: carat = -1
        return (-carat, r.get('_source_row', 0))
    records.sort(key=sk)
    serial_counts = Counter(str(r.get('Serial No')) for r in records if not is_blank(r.get('Serial No')))
    dups = sum(c - 1 for c in serial_counts.values() if c > 1)
    if dups:
        add_issue(audit, dataset, 'WARNING', 'Potential duplicate Serial No kept', source_label(raw_path),
                  details=f'{dups} repeated Serial No occurrence(s) found. No deduplication is applied.')
    write_stock_output(out_name, headers, records)
    audit_blanks(audit, dataset, headers, records)
    status = 'PASS' if all(m['Status'] != 'MISSING' or m['Final Column'] in DERIVED_COLUMNS for m in mapping_rows) else 'FAILED'
    summary = {
        'Dataset': dataset, 'Output File': out_name, 'Raw File': source_label(raw_path),
        'Previous Final File': '', 'Source Rule': 'RAW Stock only; no previous Stock file required; no history merge; no deduplication',
        'Window Rule': 'Current snapshot', 'Date Field': '', 'Processing Date': processing_date().isoformat(), 'Anchor Date': '', 'Cutoff Date': '',
        'History Rows': 'Not used', 'History Invalid Dates': 'Not used', 'Latest History Timestamp': 'Not used',
        'Raw Rows Read': raw_rows, 'Raw Valid Dates': '', 'Raw Invalid Dates': 0, 'Raw New Rows Appended': 'All valid raw rows are source',
        'Raw Historical/Overlap Rows Skipped': 0, 'Blacklist Removed From Raw': 0, 'Stock Non-Data Rows Removed': nondata_rows_removed,
        'Before Cutoff': 0, 'After Anchor': 0,
        'Output Rows': len(records), 'Potential Duplicates Kept': dups, 'Status': status
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
    audit = {'summary': [], 'mapping': [], 'blanks': [], 'issues': []}
    ok_flags = []
    cart_raw = locate([r'ADD[_ -]?TO[_ -]?CART', r'Cart'], RAW_DIR)
    hold_raw = locate([r'JANGAD[_ -]?HISTORY', r'^JANGAD$', r'Hold'], RAW_DIR)
    offer_raw = locate([r'OFFER[_ -]?PLACE', r'Offer'], RAW_DIR)
    view_raw = locate([r'APPOINTMENT', r'View'], RAW_DIR)
    wishlist_raw = locate([r'WISHLIST', r'Wish'], RAW_DIR)
    sales_raw = locate_sales(RAW_DIR)
    stock_raw = locate([r'^Stock$', r'^KSTAKA', r'^KSSTAKA'], RAW_DIR)

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
            'Cart/Offer/Wishlist': 'Previous final preserved before its last calendar date; if RAW contains that date, replace the entire boundary date from RAW; otherwise append after latest timestamp; 90-day window; exact-row dedup deferred',
            'View': 'Previous final preserved before boundary date; boundary date refreshed from RAW; STONE_ID kept as one SERIAL_NO value; blank/0 removed; processing-date 90-day window; exact-row dedup deferred',
            'Hold': 'Same whole-boundary-date refresh rule; SERIAL_NO comes from RAW SERIAL_NO; 90-day window from processing date; exact-row dedup deferred',
            'Sales': 'Sales 120/180 input files carry descriptive history. Current RAW updates SPer/SRate on exact Serial No + Sale Date overlaps and appends new transactions. Blank Lab/Flo become NON.',
            'Stock': 'RAW Stock only; no previous Stock file required; canonical output headers are built into the processor; no deduplication',
            'Derived Fields': 'Sales Carat_Grp/Amt and Stock Carat_Grp/Tag are computed in Python and written as values; blank Lab/Flo become NON; downstream logic does not depend on Excel cached formulas.',
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
