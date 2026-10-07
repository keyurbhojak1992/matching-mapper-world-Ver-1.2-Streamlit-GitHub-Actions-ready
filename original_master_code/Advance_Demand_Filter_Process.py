import pandas as pd
import zipfile
import os
from glob import glob
from io import BytesIO
import re
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Border, Side, PatternFill, Font, numbers
from openpyxl import load_workbook, Workbook
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor, TwoCellAnchor
from openpyxl.worksheet.hyperlink import Hyperlink
from copy import copy
import tempfile
import shutil

# --- USER INPUT SECTION ---
# ==========================
# Set the tolerance level for matching:
# 1: for a +/-1 tolerance (Recommended)
TOLERANCE_LEVEL = 1

# --- FANCY COLOR TOLERANCE CONTROL ---
# False: Fancy Color uses EXACT matching only (ignores +/- TOLERANCE_LEVEL).
# True:  Fancy Color uses TOLERANCE_LEVEL against FANCY_COLOR_ORDER.
# Example with TOLERANCE_LEVEL = 1 and filter PINK:
#   False -> PINK only
#   True  -> YELLOW, PINK, BLUE
ENABLE_FANCY_COLOR_TOLERANCE = False

# --- FLUORESCENCE CONTROL ---
# True: Groups VSL, SL, FNT together AND makes NON next to FNT.
#       Request 'NON' -> Matches NON + FNT (VSL/SL).
# False: Strict distinct order (NON -> VSL -> SL -> FNT).
ENABLE_FLO_MAPPING = False

# --- RELATED TAG CONTROL ---
# True: Adds the 'Tag_No' column to the output file showing the tag from the filters file.
# False: Standard behavior, no new column added.
ENABLE_RELATED_TAG = False

# --- NO NEED / CLEAN FILE EXCLUSION CONTROL ---
# True: Apply exclusion rules from no_need.xlsx (same logic as Regular Suggested code).
# False: Ignore no_need.xlsx and keep the original/base output.
ENABLE_CLEAN_FILE = True

# --- OUTPUT FILTER CONTROL NEW ---
# "ALL": Apply the allowed Match Status list to output rows.
# "NONE": Disable this output-only filter and keep the original/base output.
MATCH_STATUS_FILTER_MODE = "NONE"  # "ALL" / "NONE"
MATCH_STATUS_ALLOWED_TYPES = ["Exact", "Near by"]

# ["NONE"] = apply to all parties when MODE is "ALL".
# Put party names here to apply the output status filter only to those parties.
# Example: ["ABC Diamonds", "XYZ Gems"]
MATCH_STATUS_SPECIFIC_PARTIES = ["NONE"]
# ==========================

COLOR_ORDER = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O-P', 'Q-R', 'S-T', 'U-V', 'W-X', 'Y-Z']

# Fancy Color order supplied for Filters.xlsx ordered matching.
# ENABLE_FANCY_COLOR_TOLERANCE controls whether +/- TOLERANCE_LEVEL is applied.
FANCY_COLOR_ORDER = [
    'YELLOW', 'PINK', 'BLUE', 'RED', 'GREEN', 'PURPLE', 'ORANGE', 'VIOLET', 'GRAY', 'BLACK', 'BROWN',
    'BROWN-PINK', 'BROWN-YELLOW', 'BROWN-PURPLE', 'ORANGE-BROWN', 'PURPLE-PINK', 'PINK-PURPLE',
    'ORANGE-YELLOW', 'PINK-BROWN', 'GREEN-YELLOW', 'BROWN-ORANGE', 'BROWN-GREEN-YELLOW',
    'GRAY-GREEN', 'BROWN-GREEN', 'O-P', 'Q-R', 'S-T', 'U-V', 'W-X', 'Y-Z',
    'YELLOW-GREEN', 'YELLOW-GRAY', 'GRAY-BLUE', 'YELLOW-BROWN', 'YELLOW-ORANGE',
    'GRAY-PURPLE', 'PURPLE-RED', 'ORANGE-PINK', 'GREEN-GRAY'
]

CLARITY_ORDER = ['FL', 'IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2', 'SI3', 'I1', 'I2', 'I3', 'I4', 'I5']

# --- DYNAMIC FLUORESCENCE ORDER ---
if ENABLE_FLO_MAPPING:
    FLUORESCENCE_ORDER = ['NON', 'FNT', 'MED', 'STG', 'VST']
else:
    FLUORESCENCE_ORDER = ['NON', 'VSL', 'SL', 'FNT', 'MED', 'STG', 'VST']

# New order lists for advanced filtering
CUT_POL_SYM_ORDER = ["F", "G", "VG", "EX"]
SHADE_ORDER = ['NONE', 'INSPECTION ADVISABLE', 'MIX TINGE 1', 'MIX TINGE 2', 'MIX TINGE 3', 'LIGHT BROWN', 'BROWN', 'STRONG BROWN', 'VERY STRONG BROWN']
MILKY_ORDER = ['NONE', 'LIGHT MILKY', 'LIGHT CLOUD', 'MILKY', 'STRONG MILKY', 'VERY STRONG MILKY', 'INSPECTION ADVISABLE']
E_FACET_ORDER = ['NON', 'PS', 'CS', 'PM', 'CM', 'TEF']
TOIN_SOIN_GIRDLE_PAV_OPEN_ORDER = ['NONE', 'HAIR LINE', 'SMALL OPEN', 'MEDIUM OPEN', 'HEAVY OPEN', 'HVY OPEN']
TBINC_SBINC_ORDER = ['NONE', 'PIN POINT BLK', 'MINOR BLK', 'NOTICABLE BLACK', 'MEDIUM BLK', 'HEAVY BLK']

# New order list for sorting
SHAPE_ORDER = [
    "RBC", "PEAR", "PMB", "OVAL", "OMB", "MQ", "MMB", "EME", "SEM", "RMB", "RM", "ECB", "ECMB", "ECM", "ECBF", "CB", "CMB", "CM", "HRT", "HMB", "PR", "TR", "BGT", "PMC", "OMC", "MMC", "SRMB", "SRM", "TP", "TPB", "SLD", "CAD", "SCAD", "TB", "LOZG", "FLM", "HFMS", "HFM", "KITE", "KMSC", "HEXA", "HRS", "BUT", "BULL", "GANE", "GUN", "CU", "RAD", "DECA", "UNQ", "ECA", "CMC", "OLMC", "OEC", "HSC", "BULT", "SSC", "MOVL", "PSC", "RMC", "ORC", "RRC", "OLMB"
]

ORDER_MAPPING = {
    'COLOR NAME': COLOR_ORDER,
    'FANCY COLOR': FANCY_COLOR_ORDER,
    'QUA': CLARITY_ORDER,
    'FLO': FLUORESCENCE_ORDER,
    'CUT': CUT_POL_SYM_ORDER,
    'POL': CUT_POL_SYM_ORDER,
    'SYM': CUT_POL_SYM_ORDER,
    'SHADE': SHADE_ORDER,
    'MILKY': MILKY_ORDER,
    'E.FACET': E_FACET_ORDER,
    'TOIN_NAME': TOIN_SOIN_GIRDLE_PAV_OPEN_ORDER,
    'SOIN_NAME': TOIN_SOIN_GIRDLE_PAV_OPEN_ORDER,
    'GIRDLE_OPEN_NAME': TOIN_SOIN_GIRDLE_PAV_OPEN_ORDER,
    'PAV_OPEN_NAME': TOIN_SOIN_GIRDLE_PAV_OPEN_ORDER,
    'TBINC_NAME': TBINC_SBINC_ORDER,
    'SBINC_NAME': TBINC_SBINC_ORDER,
    'TABLE INCL': [],
}

DECIMAL_COLUMNS = ['Carat', 'RapRate', 'Reserve%', 'Rate', 'Amount', 'Ratio', 'Dep', 'Tab', 'Girdle', 'C/A', 'C/H', 'P/A', 'P/H', 'Disc%', 'Price/carat', 'Total Amount', 'Depth', 'Table']

#DESIRED_COLUMN_SEQUENCE = [
#    "Sr.No", "Tag_No", "Serial No", "DETAIL", "Shp", "Carat", "Color", "Qua", "Cut", "Pol", "Sym",
#    "Flo", "Lab", "ReportNo", "RapRate", "Reserve%", "Disc%", "Rate", "Amount", "Measurement",
#    "Dep", "Tab", "Ratio", "Shade", "Milky", "E.Facet", "TOIN_NAME", "SOIN_NAME", "GIRDLE_OPEN_NAME", "PAV_OPEN_NAME", "TBINC_NAME", "SBINC_NAME", "Remarks"
#]

DESIRED_COLUMN_SEQUENCE = [
    "Sr.No", "Tag_No", "Serial No", "DETAIL", "PHYGITAL", "Shp", "Carat", "Color", "Qua", "Cut", "Pol", "Sym",
    "Flo", "Lab", "ReportNo", "RapRate", "Reserve%", "Disc%", "Rate", "Amount", "SGPComment", "Loc", "Measurement",
    "Ratio", "Dep", "Tab", "Dia", "Girdle", "C/A", "C/H", "P/A", "P/H", "EYE Clean", "Luster", "E.Facet", "Culet", "TINC_NAME",
    "SINC_NAME", "TOIN_NAME", "SOIN_NAME", "GIRDLE_OPEN_NAME", "PAV_OPEN_NAME", "TBINC_NAME", "SBINC_NAME",
    "Spot Incl.", "KeyToSymbol", "Comment", "Remarks", "IMAGE", "VIDEO", "CERTIFICATE", "Shade", "Milky"]

STOCK_COLUMN_MAPPING = {
    "Loc": ["Location"],
    "Sr.NO": ["Sr.NO"],
    "Serial No": ["Serial No"],
    "DETAIL": ["DETAIL"],
    "Shp": ["Shp", "Shape"],
    "Carat": ["Carat"],
    "Color": ["Color name", "Color"],
    "Qua": ["Qua", "Clarity"],
    "Cut": ["Cut", "cut"],
    "Pol": ["Pol", "polish"],
    "Sym": ["Sym"],
    "Flo": ["Flo", "Flou."],
    "Lab": ["Lab"],
    "RapRate": ["RapRate", "Rap"],
    "Disc%": ["Reserve%", "Disc%", "Dis%"],
    "Rate": ["Rate", "Price/Ct$"],
    "Amount": ["Amount", "Amount$"],
    "Measurement": ["Measurement", "Measurment"], # 👈 ADDED: Fixes the alignment and typo gap
    "Milky": ["Milky"],
    "Shade": ["Shade"],
    "L": ["L", "Length"],
    "W": ["W", "Width"],
    "H": ["H", "Height"],
    "Dep": ["Dep", "Depth%"],
    "Tab": ["Tab", "Table%"],
    "Ratio": ["Ratio"],
    "KeyToSymbol": ["KeyToSymbol", "Key To Symbol"],
    "comment": ["comment", "Comment"],
    "TINC_NAME": ["Table Incl", "Table Inclusion"],
    "TOIN_NAME": ["TOIN_NAME", "Table Open"],
    "TBINC_NAME": ["TBINC_NAME", "Table Black"],
    "SINC_NAME": ["SINC_NAME", "Side Inclusion"],
    "SOIN_NAME": ["SOIN_NAME", "Side Open"],
    "SBINC_NAME": ["SBINC_NAME", "Side Black"],
    "GIRDLE_OPEN_NAME": ["GIRDLE_OPEN_NAME", "Girdle Open"],
    "C/A": ["C/A", "CR Angle"],
    "C/H": ["C/H", "CR Height"],
    "P/A": ["P/A", "Pav Angle"],
    "P/H": ["P/H", "Pav Height"],
    "Culet": ["Culet"],
    "Girdle": ["Girdle", "Girdle %"],
    "IMAGE": ["IMAGE"],
    "VIDEO": ["VIDEO"],
    "CERTIFICATE": ["CERTIFICATE"],
    "PHYGITAL": ["PHYGITAL"],
    "ReportNo": ["ReportNo", "Report No", "REPORT NO"], # 👈 ADDED: Catches the space format
    "SGPComment": ["SGPComment", "SGP Comment", "SGP COMMENT"],
    "Fancy Color": ["Fancy Color", "FancyColor", "Fancy Colour", "FancyColour"],
    "Dia": ["Dia"],
    "EYE Clean": ["EYE Clean"],
    "Luster": ["Luster"],
    "E.Facet": ["E.Facet"],
    "ShapeDesc": ["ShapeDesc"],
    "SUB_LOCATION": ["SUB_LOCATION"]
}

# --- FUTURE-PROOF COLUMN / TEXT MATCHING HELPERS ---
# These helpers let new filter/no_need columns automatically match a stock column
# when the header is the same apart from spaces, punctuation, or letter case.
METADATA_FILTER_KEYS = {
    'SOLDPARTY', 'NORMALIZEDPARTY', 'SALESPERSON', 'TAGNO', 'COUNTRY', 'CITY',
    'REMARKS', 'SPECIALREMARKS'
}

# Long text fields use "contains any listed term" logic instead of exact-cell equality.
LONG_TEXT_COLUMN_KEYS = {'SGPCOMMENT', 'KEYTOSYMBOL', 'COMMENT', 'COLORDESC'}

# Known numeric fields. Future genuinely numeric stock columns are detected automatically too.
NUMERIC_FILTER_KEYS = {
    'CARAT', 'RESERVE', 'DISC', 'RATIO', 'RATE', 'RAPRATE', 'AMOUNT', 'DEP', 'TAB',
    'CA', 'CH', 'PA', 'PH', 'GIRDLE', 'DIA'
}


def normalize_column_key(name):
    """
    Normalize a header so e.g. 'SGP Comment' == 'SGPComment' and 'C/A' == 'CA'.

    Pandas automatically renames duplicate Excel headers as ``Header.1``, ``Header.2``, etc.
    Strip only that generated numeric suffix so repeated Filters.xlsx columns (currently the
    three SGP Comment columns) all resolve back to the same stock header.
    """
    clean_name = re.sub(r'\.\d+$', '', str(name).strip())
    return re.sub(r'[^A-Z0-9]+', '', clean_name.upper())


def build_column_lookup(columns):
    lookup = {}
    for col in columns:
        lookup.setdefault(normalize_column_key(col), col)
    return lookup


# Fancy stock stores the grade inside the normal Color column (for example
# "FANCY INTENSE YELLOW", "FANCY BROWNISH PINK", "S-T").  Convert those
# descriptions to the canonical Fancy Color buckets used by Filters.xlsx.
FANCY_COLOR_WORD_ALIASES = {
    'YELLOW': 'YELLOW', 'YELLOWISH': 'YELLOW',
    'PINK': 'PINK', 'PINKISH': 'PINK',
    'BLUE': 'BLUE', 'BLUISH': 'BLUE',
    'RED': 'RED', 'REDDISH': 'RED',
    'GREEN': 'GREEN', 'GREENISH': 'GREEN',
    'PURPLE': 'PURPLE', 'PURPLISH': 'PURPLE',
    'ORANGE': 'ORANGE', 'ORANGY': 'ORANGE', 'ORANGISH': 'ORANGE',
    'VIOLET': 'VIOLET', 'VIOLETISH': 'VIOLET',
    'GRAY': 'GRAY', 'GREY': 'GRAY', 'GRAYISH': 'GRAY', 'GREYISH': 'GRAY',
    'BLACK': 'BLACK', 'BLACKISH': 'BLACK',
    'BROWN': 'BROWN', 'BROWNISH': 'BROWN',
}

FANCY_COLOR_DESCRIPTOR_WORDS = {
    'FANCY', 'FAINT', 'VERY', 'LIGHT', 'INTENSE', 'VIVID', 'DEEP', 'DARK'
}


def normalize_fancy_color_value(value):
    """Return the stock Color value as one of the configured Fancy Color buckets.

    Examples:
      FANCY INTENSE YELLOW        -> YELLOW
      FANCY BROWNISH PINK         -> BROWN-PINK
      FANCY BROWNISH GREENISH YELLOW -> BROWN-GREEN-YELLOW
      FANCY GRAYISH PINK-PURPLE   -> PINK-PURPLE
      S-T                         -> S-T

    When a stock description contains more color modifiers than the configured
    buckets support, prefer the longest right-most supported combination.  This
    keeps the dominant/right-most color wording intact without inventing a new bucket.
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ''

    text = str(value).strip().upper()
    if not text:
        return ''

    # Color Desc may contain suffixes such as ", NATURAL EVEN"; only the color grade matters.
    text = text.split(',', 1)[0].strip()
    text = re.sub(r'\[\*\]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()

    # Preserve the configured alphabetic ranges exactly (O-P through Y-Z), and
    # also accept a stock column that already contains a canonical Fancy Color value.
    if text in FANCY_COLOR_ORDER:
        return text

    raw_words = re.findall(r'[A-Z]+', text)
    colors = []
    for word in raw_words:
        if word in FANCY_COLOR_DESCRIPTOR_WORDS:
            continue
        mapped = FANCY_COLOR_WORD_ALIASES.get(word)
        if mapped and (not colors or colors[-1] != mapped):
            colors.append(mapped)

    if not colors:
        return ''

    # First try the complete color phrase (needed for BROWN-GREEN-YELLOW).
    full = '-'.join(colors)
    if full in FANCY_COLOR_ORDER:
        return full

    # If the stock phrase has an extra modifier not represented in the allowed list,
    # take the longest right-most contiguous supported combination.
    max_len = min(3, len(colors))
    for size in range(max_len, 1, -1):
        for start in range(len(colors) - size, -1, -1):
            candidate = '-'.join(colors[start:start + size])
            if candidate in FANCY_COLOR_ORDER:
                return candidate

    # Finally fall back to the dominant/right-most single color when supported.
    for color in reversed(colors):
        if color in FANCY_COLOR_ORDER:
            return color

    return ''


def resolve_stock_column(filter_col, df, explicit_mapping=None):
    """Resolve known aliases first, then dynamically match future columns by normalized header."""
    explicit_mapping = explicit_mapping or {}
    mapped = explicit_mapping.get(filter_col)
    if mapped in df.columns:
        return mapped
    return build_column_lookup(df.columns).get(normalize_column_key(filter_col))


def normalize_compact_text(value):
    """
    Normalize text for flexible word/phrase matching.
    Examples: 'Pin Point', 'PIN-POINT', 'pin_point', 'Pinpoint' -> 'PINPOINT'.
    """
    if value is None:
        return ''
    return re.sub(r'[^A-Z0-9]+', '', str(value).upper())


def split_rule_terms(value):
    """Comma-separated rule terms; empty terms are ignored."""
    return [v.strip() for v in str(value).split(',') if str(v).strip()]


def split_long_text_include_exclude_terms(value):
    """
    Regular Suggested-compatible long-text syntax:
    - ``CRYSTAL,NEEDLE``       -> include CRYSTAL OR NEEDLE
    - ``!FEATHER,!CLOUD``      -> exclude FEATHER AND CLOUD matches
    - ``CRYSTAL,!FEATHER``     -> must contain CRYSTAL and must not contain FEATHER

    This is used for KeyToSymbol / SGP Comment / Comment / Color Desc.
    """
    include_terms = []
    exclude_terms = []
    for raw_term in split_rule_terms(value):
        term = raw_term.strip()
        if term.startswith('!'):
            term = term[1:].strip()
            if term:
                exclude_terms.append(term)
        elif term:
            include_terms.append(term)
    return include_terms, exclude_terms


def contains_any_normalized(series, terms):
    """
    True when a stock text cell contains ANY listed term after ignoring case,
    spaces and punctuation. Extra words in the stock cell are allowed.
    """
    normalized_terms = [normalize_compact_text(v) for v in terms]
    normalized_terms = [v for v in normalized_terms if v]
    if not normalized_terms:
        return pd.Series(False, index=series.index)

    normalized_series = series.fillna('').astype(str).map(normalize_compact_text)
    mask = pd.Series(False, index=series.index)
    for term in normalized_terms:
        mask |= normalized_series.str.contains(re.escape(term), regex=True, na=False)
    return mask


def build_long_text_filter_condition(series, filter_value):
    """Build the normal Filters.xlsx condition for long-text columns."""
    include_terms, exclude_terms = split_long_text_include_exclude_terms(filter_value)
    condition = pd.Series(True, index=series.index)
    has_rule = False

    if include_terms:
        # Any positive term can match.
        condition &= contains_any_normalized(series, include_terms)
        has_rule = True

    if exclude_terms:
        # No excluded term may appear.
        condition &= ~contains_any_normalized(series, exclude_terms)
        has_rule = True

    return condition if has_rule else pd.Series(False, index=series.index)


def is_numeric_stock_column(series):
    """Detect true numeric columns while avoiding text columns that merely contain some numbers."""
    if pd.api.types.is_numeric_dtype(series):
        return True
    nonblank = series.fillna('').astype(str).str.strip()
    nonblank = nonblank[nonblank != '']
    if nonblank.empty:
        return False
    parsed = pd.to_numeric(nonblank, errors='coerce')
    return bool(parsed.notna().mean() >= 0.95)

def get_carat_rank(carat_value):
    if not isinstance(carat_value, (int, float)):
        return 99

    if carat_value >= 15.00: return 1
    elif 8.00 <= carat_value <= 14.99: return 2
    elif 5.00 <= carat_value <= 7.99: return 3
    elif 4.50 <= carat_value <= 4.99: return 4
    elif 4.00 <= carat_value <= 4.49: return 5
    elif 3.50 <= carat_value <= 3.99: return 6
    elif 3.00 <= carat_value <= 3.49: return 7
    elif 2.50 <= carat_value <= 2.99: return 8
    elif 2.00 <= carat_value <= 2.49: return 9
    elif 1.70 <= carat_value <= 1.99: return 10
    elif 1.50 <= carat_value <= 1.69: return 11
    elif 1.25 <= carat_value <= 1.49: return 12
    elif 1.20 <= carat_value <= 1.24: return 13
    elif 1.10 <= carat_value <= 1.19: return 14
    elif 1.00 <= carat_value <= 1.09: return 15
    elif 0.90 <= carat_value <= 0.99: return 16
    elif 0.80 <= carat_value <= 0.89: return 17
    elif 0.70 <= carat_value <= 0.79: return 18
    elif 0.60 <= carat_value <= 0.69: return 19
    elif 0.50 <= carat_value <= 0.59: return 20
    elif 0.40 <= carat_value <= 0.49: return 21
    elif 0.30 <= carat_value <= 0.39: return 22

    return 99

def format_native_hyperlinks(worksheet, dataframe):
    """Bypasses Excel's 255-char formula limit by applying native XML hyperlinks."""
    from openpyxl.styles import Font # Ensuring Font is available
    link_columns = ["CERTIFICATE", "VIDEO", "IMAGE", "DETAIL", "PHYGITAL", "LAYOUT"]
    blue_font = Font(color="0000FF", underline="single")

    # Find which link columns actually exist in this specific dataframe
    col_indices = {}
    for col_name in link_columns:
        if col_name in dataframe.columns:
            col_indices[col_name] = dataframe.columns.get_loc(col_name) + 1 # 1-based for openpyxl

    if not col_indices:
        return

    # Loop through the rows and apply the native hyperlink property
    for row_idx, row in enumerate(dataframe.itertuples(index=False), start=2): # Starts at row 2 (below header)
        for col_name, col_idx in col_indices.items():
            val = str(row[col_idx - 1])
            if 'http' in val.lower():
                cell = worksheet.cell(row=row_idx, column=col_idx)
                cell.value = col_name   # Displays "VIDEO", "CERTIFICATE", etc.
                cell.hyperlink = val    # Attaches the 400+ char URL natively
                cell.font = blue_font   # Makes it look like a clickable link

def normalize_party_name(name):
    if not isinstance(name, str): return ""
    name = re.sub(r'[^\w\s]', '', name).upper().strip()
    name = re.sub(r'\b(INC|LLC|LTD|CO|CORP|COMPANY)\b', '', name).strip()
    return name

def normalize_match_status_name(status):
    """Map user-facing output filter names to the existing internal Match_Type names."""
    status_key = re.sub(r'[^A-Z0-9]+', '', str(status).strip().upper())
    status_aliases = {
        'EXACT': 'Exact',
        'NEARBY': 'Nearby',
        'FILTER': 'Maybe',
        'MAYBE': 'Maybe',
    }
    return status_aliases.get(status_key)

def should_apply_match_status_filter(party_display_name):
    """Return True only when the new output-only status filter should affect this party."""
    if str(MATCH_STATUS_FILTER_MODE).strip().upper() != 'ALL':
        return False

    specific_parties = [
        normalize_party_name(p)
        for p in MATCH_STATUS_SPECIFIC_PARTIES
        if str(p).strip().upper() not in ('', 'NONE')
    ]

    # No actual party names supplied means ALL parties.
    if not specific_parties:
        return True

    return normalize_party_name(party_display_name) in specific_parties

def apply_match_status_output_filter(df, party_display_name):
    """Filter final output rows only; does not alter matching or tolerance logic."""
    if df.empty or 'Match_Type' not in df.columns:
        return df

    if not should_apply_match_status_filter(party_display_name):
        return df

    allowed_types = {
        normalized
        for normalized in (normalize_match_status_name(v) for v in MATCH_STATUS_ALLOWED_TYPES)
        if normalized is not None
    }

    # If the mode is enabled but no valid types are listed, keep no output rows.
    if not allowed_types:
        return df.iloc[0:0].copy()

    match_values = df['Match_Type'].astype(object).astype(str)
    return df[match_values.isin(allowed_types)].copy()

def get_tolerance_values(value, order, tolerance_level=1):
    try:
        index = order.index(value)
        start = max(0, index - tolerance_level)
        end = min(len(order) - 1, index + tolerance_level)
        return order[start:end + 1]
    except ValueError:
        return [value]

def expand_values_with_tolerance(values, order_list, tolerance_level):
    if tolerance_level == 0 or not order_list:
        return values
    expanded_values = []
    for val in values:
        if val in order_list:
            tolerance_values = get_tolerance_values(val, order_list, tolerance_level)
            expanded_values.extend(tolerance_values)
        else:
            expanded_values.append(val)
    return list(set(expanded_values))

def parse_filter_value(filter_string, order_list=None, tolerance_level=0, apply_flo_mapping=False):
    if pd.isna(filter_string) or not isinstance(filter_string, str):
        return []

    filter_string = filter_string.strip().upper()

    if apply_flo_mapping and order_list == FLUORESCENCE_ORDER:
        filter_string = filter_string.replace('VSL', 'FNT').replace('SL', 'FNT')

    values = []
    if order_list and 'TO' in filter_string:
        parts = filter_string.split('TO')
        if len(parts) == 2:
            start_val = parts[0].strip()
            end_val = parts[1].strip()
            try:
                start_idx = order_list.index(start_val)
                end_idx = order_list.index(end_val)
                if start_idx <= end_idx: values.extend(order_list[start_idx:end_idx + 1])
                else: values.extend(order_list[end_idx:start_idx + 1])
            except ValueError: values.append(filter_string)
        else: values.append(filter_string)
    elif ',' in filter_string:
        for val in filter_string.split(','):
            val_stripped = val.strip()
            if order_list and val_stripped in order_list: values.append(val_stripped)
            elif not order_list: values.append(val_stripped)
            elif val_stripped not in order_list: values.append(val_stripped)
    else:
        if order_list and filter_string in order_list: values.append(filter_string)
        else: values.append(filter_string)

    if tolerance_level > 0 and order_list:
        values = expand_values_with_tolerance(values, order_list, tolerance_level)

    return list(set(values))

def parse_numeric_filter(series, filter_string):
    filter_string = str(filter_string).strip().upper()
    try:
        if 'TO' in filter_string:
            parts = filter_string.split('TO')
            if len(parts) == 2:
                return (series >= float(parts[0].strip())) & (series <= float(parts[1].strip()))
        elif filter_string.startswith('>='): return series >= float(filter_string[2:].strip())
        elif filter_string.startswith('>'): return series > float(filter_string[1:].strip())
        elif filter_string.startswith('<='): return series <= float(filter_string[2:].strip())
        elif filter_string.startswith('<'): return series < float(filter_string[1:].strip())
        elif filter_string.startswith('!='): return series != float(filter_string[2:].strip())
        else: return series == float(filter_string)
    except: return pd.Series(False, index=series.index)

def parse_carat_filter(series, filter_string):
    filter_string = str(filter_string).strip().upper()
    base_condition = pd.Series(True, index=series.index)
    if "NO DOT" in filter_string:
        base_condition = base_condition & (~series.astype(str).str.endswith('.0', na=False))
        filter_string = filter_string.replace("NO DOT", "").strip()
    if filter_string:
        combined = pd.Series(False, index=series.index)
        for ind_filter in [f.strip() for f in filter_string.split(',')]:
            combined = combined | parse_numeric_filter(series, ind_filter)
        base_condition = base_condition & combined
    return base_condition

def parse_measurement_filter(series, filter_string):
    filter_string = str(filter_string).strip().upper()
    try:
        operator_match = re.match(r'([<>!=]+)\s*(\d+\.?\d*)', filter_string)
        if not operator_match: return pd.Series(True, index=series.index)
        operator = operator_match.group(1)
        target = float(operator_match.group(2))
        def get_dim(s):
            if pd.isna(s): return None
            parts = str(s).split('*')
            if parts:
                try: return float(parts[-1].strip())
                except: return None
            return None
        last_dims = series.apply(get_dim)
        mask = last_dims.notna()
        if operator == '>=': cond = last_dims[mask] >= target
        elif operator == '>': cond = last_dims[mask] > target
        elif operator == '<=': cond = last_dims[mask] <= target
        elif operator == '<': cond = last_dims[mask] < target
        elif operator == '!=': cond = last_dims[mask] != target
        elif operator == '=': cond = last_dims[mask] == target
        else: return pd.Series(True, index=series.index)
        full_cond = pd.Series(False, index=series.index)
        full_cond[mask] = cond
        return full_cond
    except: return pd.Series(True, index=series.index)

def build_stock_filter_cache(stocks_df):
    """Precompute reusable Series once so every party rule does not re-normalize full stock."""
    upper_cache = {}
    numeric_cache = {}

    # Columns repeatedly compared as text across party rules.
    text_cols = {
        'Shp', 'Qua', 'Cut', 'Pol', 'Sym', 'Flo', 'Lab', 'Shade', 'Milky', 'E.Facet',
        'TOIN_NAME', 'SOIN_NAME', 'GIRDLE_OPEN_NAME', 'PAV_OPEN_NAME', 'TBINC_NAME',
        'SBINC_NAME', 'Color', 'Table Incl', 'Remarks'
    }
    for col in text_cols:
        if col in stocks_df.columns:
            upper_cache[col] = stocks_df[col].astype(str).str.upper()

    # Numeric conversion is expensive when repeated for every party/filter row.
    numeric_cols = {'Disc%', 'Ratio', 'Rate', 'RapRate', 'Amount', 'Dep', 'Tab'}
    for col in numeric_cols:
        if col in stocks_df.columns:
            numeric_cache[col] = pd.to_numeric(stocks_df[col], errors='coerce')

    measurement_last = None
    if 'Measurement' in stocks_df.columns:
        # Equivalent to the existing logic: use the final *-separated dimension.
        measurement_last = pd.to_numeric(
            stocks_df['Measurement'].astype(str).str.rsplit('*', n=1).str[-1].str.strip(),
            errors='coerce'
        )

    return {
        'upper': upper_cache,
        'numeric': numeric_cache,
        'measurement_last': measurement_last,
        'index': stocks_df.index,
    }


def parse_measurement_filter_cached(last_dims, filter_string):
    if last_dims is None:
        return None
    filter_string = str(filter_string).strip().upper()
    try:
        operator_match = re.match(r'([<>!=]+)\s*(\d+\.?\d*)', filter_string)
        if not operator_match:
            return pd.Series(True, index=last_dims.index)
        operator = operator_match.group(1)
        target = float(operator_match.group(2))
        if operator == '>=': return last_dims >= target
        if operator == '>': return last_dims > target
        if operator == '<=': return last_dims <= target
        if operator == '<': return last_dims < target
        if operator == '!=': return last_dims != target
        if operator == '=': return last_dims == target
        return pd.Series(True, index=last_dims.index)
    except Exception:
        return pd.Series(True, index=last_dims.index)


def apply_tolerance_filter(df, filter_rules, tolerance_level, stock_cache=None):
    """
    Apply the normal Filters.xlsx rule.

    Future-proof behavior:
    - Known business columns keep their existing special/tolerance logic.
    - A new filter column is automatically used when a matching stock header exists.
    - Long-text columns (KeyToSymbol, SGP Comment, Comment, Color Desc) support OR includes
      plus !term exclusions, following the Regular Suggested filter behavior.
    - Repeated Excel headers such as SGP Comment / SGP Comment.1 / SGP Comment.2 all map
      to the same stock column; each populated repeated column becomes an additional AND condition.
    - Fancy Color uses the supplied ordered list and derives its bucket from stock Color when needed.
    - Future numeric stock columns support numeric operators/ranges automatically.
    """
    if df.empty or filter_rules.empty:
        return pd.DataFrame()

    conditions = []
    numeric_cols_touched = set()

    column_mapping = {
        'SHP': 'Shp', 'QUA': 'Qua', 'CUT': 'Cut', 'POL': 'Pol', 'SYM': 'Sym',
        'FLO': 'Flo', 'LAB': 'Lab', 'DISC%': 'Disc%', 'SHADE': 'Shade',
        'MILKY': 'Milky', 'E.FACET': 'E.Facet', 'TOIN_NAME': 'TOIN_NAME',
        'SOIN_NAME': 'SOIN_NAME', 'GIRDLE_OPEN_NAME': 'GIRDLE_OPEN_NAME',
        'PAV_OPEN_NAME': 'PAV_OPEN_NAME', 'TBINC_NAME': 'TBINC_NAME',
        'SBINC_NAME': 'SBINC_NAME', 'MEASUREMENT': 'Measurement',
        'REMARKS': 'Remarks', 'COLOR NAME': 'Color', 'CARAT': 'Carat',
        'RESERVE%': 'Disc%', 'RATIO': 'Ratio', 'RATE': 'Rate', 'RAPRATE': 'RapRate',
        'AMOUNT': 'Amount', 'TABLE INCL': 'Table Incl', 'DEP': 'Dep', 'TAB': 'Tab',
        'C/A': 'C/A', 'C/H': 'C/H', 'P/A': 'P/A', 'P/H': 'P/H', 'GIRDLE': 'Girdle',
        'DIA': 'Dia', 'SGP COMMENT': 'SGPComment', 'SGPCOMMENT': 'SGPComment',
        'KEYTOSYMBOL': 'KeyToSymbol', 'COMMENT': 'comment', 'COLOR DESC': 'Color Desc',
        'FANCY COLOR': 'Fancy Color', 'FANCYCOLOR': 'Fancy Color',
        'EYE CLEAN': 'EYE Clean', 'LUSTER': 'Luster', 'CULET': 'Culet'
    }
    remarks = str(filter_rules.get('REMARKS', 'INCLUDE')).strip().upper()

    upper_cache = stock_cache.get('upper', {}) if stock_cache else {}
    numeric_cache = stock_cache.get('numeric', {}) if stock_cache else {}

    for filter_col, filter_value in filter_rules.items():
        filter_key = normalize_column_key(filter_col)
        if filter_key in METADATA_FILTER_KEYS or pd.isna(filter_value) or str(filter_value).strip() == '':
            continue

        # Fancy stock keeps its fancy grade in the normal Color column rather than
        # a literal 'Fancy Color' header. Prefer a real Fancy Color column if one exists;
        # otherwise use Color and normalize it into the configured Fancy Color buckets.
        if filter_key == 'FANCYCOLOR':
            stock_col = build_column_lookup(df.columns).get('FANCYCOLOR')
            if not stock_col and 'Color' in df.columns:
                stock_col = 'Color'
        else:
            stock_col = resolve_stock_column(filter_col, df, column_mapping)

        if not stock_col:
            # Future column is simply ignored when it does not exist in stock.
            continue

        try:
            val_str = str(filter_value).strip().upper()

            # Shape and Lab remain strict (no +/- tolerance).
            if filter_col in ['SHP', 'LAB']:
                acc = parse_filter_value(val_str, tolerance_level=0)
                series_upper = upper_cache.get(stock_col)
                if series_upper is None:
                    series_upper = df[stock_col].fillna('').astype(str).str.strip().str.upper()
                conditions.append(series_upper.isin(acc) if acc else pd.Series(False, index=df.index))

            elif filter_col == 'CARAT':
                numeric_series = pd.to_numeric(df[stock_col], errors='coerce')
                numeric_cols_touched.add(stock_col)
                conditions.append(parse_carat_filter(numeric_series, val_str))

            elif filter_col == 'MEASUREMENT':
                cached_last = stock_cache.get('measurement_last') if stock_cache else None
                if cached_last is not None and stock_col == 'Measurement':
                    conditions.append(parse_measurement_filter_cached(cached_last, val_str))
                else:
                    conditions.append(parse_measurement_filter(df[stock_col], val_str))

            elif filter_col in ORDER_MAPPING:
                order = ORDER_MAPPING[filter_col]
                apply_map = (ENABLE_FLO_MAPPING and filter_col == 'FLO')

                # Fancy Color has its own independent tolerance switch.
                # False = exact ordered value only; True = use global TOLERANCE_LEVEL.
                effective_tolerance = (
                    tolerance_level
                    if filter_key != 'FANCYCOLOR' or ENABLE_FANCY_COLOR_TOLERANCE
                    else 0
                )
                acc = parse_filter_value(
                    val_str, order, effective_tolerance, apply_flo_mapping=apply_map
                )

                if filter_key == 'FANCYCOLOR':
                    # Current fancy stock has values such as 'FANCY INTENSE YELLOW'
                    # in Color. Normalize those descriptions before ordered/tolerance matching.
                    series_to_check = df[stock_col].map(normalize_fancy_color_value)
                else:
                    series_to_check = upper_cache.get(stock_col)
                    if series_to_check is None:
                        series_to_check = df[stock_col].fillna('').astype(str).str.strip().str.upper()
                    if apply_map:
                        series_to_check = series_to_check.replace(['VSL', 'SL'], 'FNT')

                conditions.append(series_to_check.isin(acc) if acc else pd.Series(False, index=df.index))

            elif filter_key in LONG_TEXT_COLUMN_KEYS:
                # Regular Suggested behavior for KeyToSymbol / SGP Comment:
                # positives are OR, !terms are exclusions; positive and negative groups combine with AND.
                # Matching still ignores case/spaces/punctuation (Pin Point == Pinpoint == Pin-Point).
                conditions.append(build_long_text_filter_condition(df[stock_col], val_str))

            elif filter_key in NUMERIC_FILTER_KEYS or is_numeric_stock_column(df[stock_col]):
                numeric_series = numeric_cache.get(stock_col)
                if numeric_series is None:
                    numeric_series = pd.to_numeric(df[stock_col], errors='coerce')
                numeric_cols_touched.add(stock_col)
                # parse_carat_filter also supports comma-separated numeric alternatives.
                conditions.append(parse_carat_filter(numeric_series, val_str))

            else:
                # Generic future text column: comma means OR; comparison is case-insensitive exact value.
                acc = parse_filter_value(val_str, tolerance_level=0)
                series_upper = df[stock_col].fillna('').astype(str).str.strip().str.upper()
                conditions.append(series_upper.isin(acc) if acc else pd.Series(False, index=df.index))

        except Exception as exc:
            print(f"    WARNING: filter column '{filter_col}' could not be applied: {exc}")
            continue

    if not conditions:
        return pd.DataFrame()

    final_cond = conditions[0].copy()
    for c in conditions[1:]:
        final_cond &= c

    result = df.loc[~final_cond if remarks == 'EXCLUDE' else final_cond].copy()

    for col in numeric_cols_touched:
        numeric_series = numeric_cache.get(col)
        if numeric_series is None:
            numeric_series = pd.to_numeric(df[col], errors='coerce')
        result[col] = numeric_series.loc[result.index]

    return result

def classify_match_type(row, filter_rules):
    grading_cols = { 'COLOR NAME': 'Color', 'QUA': 'Qua', 'FLO': 'Flo' }
    deviation_count = 0

    for filter_col, stock_col in grading_cols.items():
        if filter_col not in filter_rules: continue
        filter_value = filter_rules[filter_col]
        if pd.isna(filter_value) or str(filter_value).strip() == '': continue

        if stock_col not in row: continue
        row_val = str(row[stock_col]).strip().upper()

        if ENABLE_FLO_MAPPING and stock_col == 'Flo':
            if row_val in ['VSL', 'SL']: row_val = 'FNT'

        val_str = str(filter_value).strip().upper()
        order = ORDER_MAPPING[filter_col]
        apply_map = (ENABLE_FLO_MAPPING and filter_col == 'FLO')

        exact_targets = parse_filter_value(val_str, order, tolerance_level=0, apply_flo_mapping=apply_map)

        if row_val not in exact_targets:
            deviation_count += 1

    if deviation_count == 0: return 'Exact'
    elif deviation_count == 1: return 'Nearby'
    else: return 'Maybe'

def classify_match_types_vectorized(df, filter_rules):
    """Vectorized equivalent of classify_match_type for Color/Clarity/Flo grading."""
    if df.empty:
        return pd.Series(dtype=object, index=df.index)

    deviation_count = pd.Series(0, index=df.index, dtype='int8')
    grading_cols = {'COLOR NAME': 'Color', 'QUA': 'Qua', 'FLO': 'Flo'}

    for filter_col, stock_col in grading_cols.items():
        if filter_col not in filter_rules:
            continue
        filter_value = filter_rules[filter_col]
        if pd.isna(filter_value) or str(filter_value).strip() == '' or stock_col not in df.columns:
            continue

        row_vals = df[stock_col].astype(str).str.strip().str.upper()
        if ENABLE_FLO_MAPPING and stock_col == 'Flo':
            row_vals = row_vals.replace(['VSL', 'SL'], 'FNT')

        val_str = str(filter_value).strip().upper()
        order = ORDER_MAPPING[filter_col]
        apply_map = (ENABLE_FLO_MAPPING and filter_col == 'FLO')
        exact_targets = parse_filter_value(val_str, order, tolerance_level=0, apply_flo_mapping=apply_map)
        deviation_count += (~row_vals.isin(exact_targets)).astype('int8')

    out = pd.Series('Maybe', index=df.index, dtype=object)
    out.loc[deviation_count == 0] = 'Exact'
    out.loc[deviation_count == 1] = 'Nearby'
    return out


def auto_adjust_column_widths(worksheet, dataframe=None):
    """Same widths as before, but use vectorized DataFrame length checks when available."""
    fixed_cols = {"CERTIFICATE", "VIDEO", "IMAGE", "DETAIL", "PHYGITAL"}

    if dataframe is not None and not dataframe.empty:
        for col_num, col_name in enumerate(dataframe.columns, 1):
            col_letter = get_column_letter(col_num)
            if col_name in fixed_cols:
                worksheet.column_dimensions[col_letter].width = 15
                continue

            header_len = len(str(col_name)) if col_name is not None else 0
            try:
                vals = dataframe[col_name]
                # Match worksheet behavior: None is ignored; other values use str(value).
                lengths = vals[vals.notna()].astype(str).str.len()
                max_data_len = int(lengths.max()) if not lengths.empty else 0
            except Exception:
                max_data_len = 0
            max_length = max(header_len, max_data_len)
            worksheet.column_dimensions[col_letter].width = (max_length + 2) * 1.2
        return

    # Navigation/small-sheet fallback.
    for column in worksheet.columns:
        col_letter = get_column_letter(column[0].column)
        col_name = column[0].value
        if col_name in fixed_cols:
            worksheet.column_dimensions[col_letter].width = 15
        else:
            max_length = max((len(str(cell.value)) for cell in column if cell.value is not None), default=0)
            worksheet.column_dimensions[col_letter].width = (max_length + 2) * 1.2


def apply_styles_to_worksheet(worksheet, dataframe=None):
    """One-pass styling equivalent to the original, reducing repeated worksheet scans."""
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")

    exact_match_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
    nearby_match_fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
    maybe_match_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    headers = [cell.value for cell in worksheet[1]]

    if worksheet.title == "Navigation":
        header_font = Font(bold=True, color="0000FF")
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
        for row in worksheet.iter_rows():
            for cell in row:
                cell.border = thin_border
                cell.alignment = center_alignment
        return

    # Header: same grey fill + border/alignment as the original.
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = center_alignment

    decimal_idx = {i for i, h in enumerate(headers) if h in DECIMAL_COLUMNS}
    sr_idx = headers.index('Sr.No') if 'Sr.No' in headers else -1
    blue_underline_font = Font(color="0000FF", underline="single")

    match_types = None
    if dataframe is not None and 'Match_Type' in dataframe.columns:
        match_types = dataframe['Match_Type'].astype(object).tolist()

    for row_offset, row_cells in enumerate(worksheet.iter_rows(min_row=2)):
        fill = None
        if match_types is not None and row_offset < len(match_types):
            mt = match_types[row_offset]
            if mt == 'Exact': fill = exact_match_fill
            elif mt == 'Nearby': fill = nearby_match_fill
            elif mt == 'Maybe': fill = maybe_match_fill

        for col_idx, cell in enumerate(row_cells):
            cell.border = thin_border
            cell.alignment = center_alignment
            if fill is not None:
                cell.fill = fill
            if col_idx in decimal_idx and isinstance(cell.value, (int, float)):
                cell.number_format = numbers.FORMAT_NUMBER_00
            if col_idx == sr_idx:
                cell.font = blue_underline_font

def process_party_data(stocks_df, party_filter_rule, tolerance_level, stock_cache=None):
    norm_party_name = party_filter_rule.get('NORMALIZED PARTY', '')
    party_display_name = party_filter_rule.get('SOLD PARTY', '')

    final_df = apply_tolerance_filter(stocks_df, party_filter_rule, tolerance_level, stock_cache=stock_cache)

    if not final_df.empty:
        final_df['Match_Type'] = classify_match_types_vectorized(final_df, party_filter_rule)
        final_df['Is_Filter_Match'] = True

        # IMPORTANT: never convert stock values to pd.Categorical directly.
        # Unknown/new values (e.g. TPB) would turn into NaN/blank. Use temporary sort ranks instead.
        shape_rank = {v: i for i, v in enumerate(SHAPE_ORDER)}
        color_rank = {v: i for i, v in enumerate(COLOR_ORDER)}
        clarity_rank = {v: i for i, v in enumerate(CLARITY_ORDER)}
        match_rank = {'Exact': 0, 'Nearby': 1, 'Maybe': 2}

        final_df['_MatchSort'] = final_df['Match_Type'].astype(str).map(match_rank).fillna(99)
        final_df['_ShapeSort'] = final_df['Shp'].fillna('').astype(str).str.upper().map(shape_rank).fillna(len(shape_rank) + 1)
        final_df['_ColorSort'] = final_df['Color'].fillna('').astype(str).str.upper().map(color_rank).fillna(len(color_rank) + 1)
        final_df['_QuaSort'] = final_df['Qua'].fillna('').astype(str).str.upper().map(clarity_rank).fillna(len(clarity_rank) + 1)
        final_df['Carat_Rank'] = final_df['Carat'].apply(get_carat_rank)

        final_df.sort_values(
            by=['_MatchSort', '_ShapeSort', 'Carat_Rank', '_ColorSort', '_QuaSort'],
            ascending=[True, True, True, True, True],
            inplace=True
        )
        final_df.drop(columns=['_MatchSort', '_ShapeSort', '_ColorSort', '_QuaSort', 'Carat_Rank'], inplace=True)

        final_df.reset_index(drop=True, inplace=True)
        if 'Sr.No' not in final_df.columns:
            final_df.insert(0, 'Sr.No', range(1, 1 + len(final_df)))

        if ENABLE_RELATED_TAG:
            tag_val = party_filter_rule.get('TAG_NO', party_filter_rule.get('TAG', ''))
            final_df['Tag_No'] = '' if pd.isna(tag_val) else tag_val

        current_cols = final_df.columns.tolist()
        ordered_cols = []

        active_sequence = DESIRED_COLUMN_SEQUENCE.copy()
        if ENABLE_RELATED_TAG and 'Tag_No' not in active_sequence:
            active_sequence.append('Tag_No')

        for col in active_sequence:
            if col in current_cols:
                ordered_cols.append(col)
        for col in current_cols:
            if col not in ordered_cols:
                ordered_cols.append(col)
        final_df = final_df[ordered_cols]

    country = party_filter_rule.get('COUNTRY', '')
    city = party_filter_rule.get('CITY', '')
    return final_df, party_display_name, country, city

def load_filter_data(filter_file_path):
    removed_parties = []
    change_format_filters = pd.DataFrame()
    try:
        xl = pd.ExcelFile(filter_file_path)
        if "Removed Party" in xl.sheet_names:
            removed_parties_df = pd.read_excel(
                filter_file_path, sheet_name="Removed Party", keep_default_na=False
            )
            if "Sold Party" in removed_parties_df.columns:
                removed_parties = removed_parties_df["Sold Party"].apply(normalize_party_name).tolist()
        if "Change Format" in xl.sheet_names:
            change_format_filters = pd.read_excel(
                filter_file_path, sheet_name="Change Format", keep_default_na=False
            )
            change_format_filters.columns = change_format_filters.columns.str.strip().str.upper()
            if "SOLD PARTY" in change_format_filters.columns and "SALES_PERSON" in change_format_filters.columns:
                change_format_filters['NORMALIZED PARTY'] = change_format_filters['SOLD PARTY'].apply(normalize_party_name)
                change_format_filters['SALES_PERSON'] = change_format_filters['SALES_PERSON'].astype(str).str.strip().str.title()
    except Exception as e:
        print(f"❌ ERROR loading filter file {filter_file_path}: {str(e)}")
    return removed_parties, change_format_filters

def prepare_change_format_filters(all_change_format_filters, removed_parties):
    """
    Base-only cleanup before processing:
    - Skip parties listed in Filters.xlsx -> Removed Party.
    - Remove duplicate filter rules.
    - When related tags are disabled, Tag_No/Tag is ignored for duplicate detection.
      If related tags are enabled later, tag-specific rules are preserved.
    """
    if all_change_format_filters.empty:
        return all_change_format_filters, 0, 0

    cleaned = all_change_format_filters.copy()

    removed_rows = 0
    if removed_parties and 'NORMALIZED PARTY' in cleaned.columns:
        removed_set = {p for p in removed_parties if str(p).strip()}
        if removed_set:
            remove_mask = cleaned['NORMALIZED PARTY'].isin(removed_set)
            removed_rows = int(remove_mask.sum())
            cleaned = cleaned.loc[~remove_mask].copy()

    ignored_for_dedup = set() if ENABLE_RELATED_TAG else {'TAG_NO', 'TAG'}
    dedup_cols = [c for c in cleaned.columns if c not in ignored_for_dedup]

    before_dedup = len(cleaned)
    if dedup_cols:
        cleaned.drop_duplicates(subset=dedup_cols, keep='first', inplace=True)
    duplicate_rows = before_dedup - len(cleaned)

    cleaned.reset_index(drop=True, inplace=True)
    return cleaned, removed_rows, duplicate_rows


def load_exclusion_data(clean_file_path):
    """Load NO_NEED rules, preserving literal text such as 'N/A'."""
    exclusion_df = pd.DataFrame()

    if not clean_file_path:
        return exclusion_df

    try:
        if not os.path.exists(clean_file_path):
            print(f"⚠️ WARNING: Clean file DOES NOT EXIST at {clean_file_path}")
            return exclusion_df

        exclusion_df = pd.read_excel(clean_file_path, keep_default_na=False)
        exclusion_df.columns = exclusion_df.columns.str.strip().str.upper()

        # Pandas names the second duplicate Remarks column REMARKS.1.
        # Treat it as a human note, not as a stock condition, so it can never conflict with REMARKS.
        rename_map = {}
        for col in exclusion_df.columns:
            if re.fullmatch(r'REMARKS\.\d+', str(col)):
                suffix = str(col).split('.', 1)[1]
                rename_map[col] = 'SPECIAL REMARKS' if suffix == '1' else f'SPECIAL REMARKS {suffix}'
        if rename_map:
            exclusion_df.rename(columns=rename_map, inplace=True)

        if "SOLD PARTY" in exclusion_df.columns:
            exclusion_df['NORMALIZED PARTY'] = exclusion_df['SOLD PARTY'].apply(normalize_party_name)

        if "SALES_PERSON" in exclusion_df.columns:
            exclusion_df['SALES_PERSON'] = exclusion_df['SALES_PERSON'].astype(str).str.strip().str.title()

    except Exception as e:
        print(f"❌ CRITICAL ERROR reading {clean_file_path}: {e}")

    return exclusion_df

def build_exclusion_rule_index(exclusion_data):
    """Index NO_NEED rules once by salesperson and party to avoid repeated DataFrame scans."""
    global_by_sp = {}
    specific_by_sp_party = {}
    if exclusion_data.empty or 'SALES_PERSON' not in exclusion_data.columns:
        return global_by_sp, specific_by_sp_party

    for salesperson, sp_df in exclusion_data.groupby('SALES_PERSON', sort=False):
        sold = sp_df['SOLD PARTY'].fillna('').astype(str).str.strip().str.upper() if 'SOLD PARTY' in sp_df.columns else pd.Series('', index=sp_df.index)
        remarks = sp_df['REMARKS'].fillna('').astype(str).str.strip().str.upper() if 'REMARKS' in sp_df.columns else pd.Series('', index=sp_df.index)
        global_mask = (sold == 'ALL') & (remarks == 'ALL')

        g = sp_df.loc[global_mask]
        if not g.empty:
            global_by_sp[salesperson] = g

        p = sp_df.loc[~global_mask]
        if not p.empty and 'NORMALIZED PARTY' in p.columns:
            for norm_party, party_df in p.groupby('NORMALIZED PARTY', sort=False):
                specific_by_sp_party[(salesperson, norm_party)] = party_df

    return global_by_sp, specific_by_sp_party


def apply_exclusion_rules(df, rules_df):
    """
    Apply NO_NEED rules.

    Important semantics:
    - Normal value (e.g. KEYTOSYMBOL='crystal,needle') means EXCLUDE rows containing ANY term.
    - Leading ! (e.g. '!crystal,needle') means KEEP rows containing ANY term and EXCLUDE
      applicable rows containing NONE of the listed terms.
    - For KeyToSymbol/other long text, spaces and punctuation are ignored, so
      Pin Point / Pinpoint / Pin-Point / PIN_POINT are treated the same.
    - New NO_NEED columns automatically work when the same stock column exists.
    """
    if df.empty or rules_df.empty:
        return df

    df_clean = df.copy()

    column_mapping = {
        'SHP': 'Shp', 'QUA': 'Qua', 'CUT': 'Cut', 'POL': 'Pol', 'SYM': 'Sym',
        'FLO': 'Flo', 'LAB': 'Lab', 'DISC%': 'Disc%', 'SHADE': 'Shade', 'MILKY': 'Milky',
        'E.FACET': 'E.Facet', 'TOIN_NAME': 'TOIN_NAME', 'SOIN_NAME': 'SOIN_NAME',
        'GIRDLE_OPEN_NAME': 'GIRDLE_OPEN_NAME', 'PAV_OPEN_NAME': 'PAV_OPEN_NAME',
        'TBINC_NAME': 'TBINC_NAME', 'SBINC_NAME': 'SBINC_NAME', 'MEASUREMENT': 'Measurement',
        'COLOR NAME': 'Color', 'CARAT': 'Carat', 'RESERVE%': 'Disc%',
        'RATIO': 'Ratio', 'RATE': 'Rate', 'RAPRATE': 'RapRate', 'AMOUNT': 'Amount',
        'TABLE INCL': 'Table Incl', 'TAB': 'Tab', 'C/A': 'C/A', 'DEP': 'Dep', 'C/H': 'C/H',
        'P/A': 'P/A', 'P/H': 'P/H', 'GIRDLE': 'Girdle', 'DIA': 'Dia',
        'SGP COMMENT': 'SGPComment', 'SGPCOMMENT': 'SGPComment',
        'KEYTOSYMBOL': 'KeyToSymbol', 'COMMENT': 'comment', 'COLOR DESC': 'Color Desc',
        'EYE CLEAN': 'EYE Clean', 'LUSTER': 'Luster', 'CULET': 'Culet'
    }

    normalized_cache = {}
    numeric_cache = {}

    def get_normalized_series(stock_col, apply_map=False):
        key = (stock_col, apply_map)
        if key not in normalized_cache:
            ser = df_clean[stock_col].fillna('').astype(str).str.replace(r'\s+', '', regex=True).str.upper()
            if apply_map:
                ser = ser.replace(['VSL', 'SL'], 'FNT')
            normalized_cache[key] = ser
        return normalized_cache[key]

    def get_numeric_series(stock_col):
        if stock_col not in numeric_cache:
            numeric_cache[stock_col] = pd.to_numeric(df_clean[stock_col], errors='coerce')
        return numeric_cache[stock_col]

    rows_to_drop_mask = pd.Series(False, index=df_clean.index)

    for _, rule in rules_df.iterrows():
        rule_mask = pd.Series(True, index=df_clean.index)
        has_criteria = False

        for filter_col, filter_value in rule.items():
            filter_key = normalize_column_key(filter_col)
            if filter_key in METADATA_FILTER_KEYS or filter_key.startswith('SPECIALREMARKS') \
               or pd.isna(filter_value) or str(filter_value).strip() == '':
                continue

            stock_col = resolve_stock_column(filter_col, df_clean, column_mapping)
            if not stock_col:
                # Human note or future column not present in current stock -> do not treat as a criterion.
                continue

            has_criteria = True
            val_str = str(filter_value).strip().upper()

            try:
                is_keep_only = val_str.startswith('!')
                clean_val = val_str[1:].strip() if is_keep_only else val_str

                if filter_col == 'MEASUREMENT':
                    match_mask = parse_measurement_filter(df_clean[stock_col], clean_val)
                    col_cond = ~match_mask if is_keep_only else match_mask

                elif filter_key in NUMERIC_FILTER_KEYS or is_numeric_stock_column(df_clean[stock_col]):
                    numeric_series = get_numeric_series(stock_col)
                    match_mask = parse_carat_filter(numeric_series, clean_val)
                    col_cond = ~match_mask if is_keep_only else match_mask

                elif filter_key in LONG_TEXT_COLUMN_KEYS:
                    terms = split_rule_terms(clean_val)
                    match_mask = contains_any_normalized(df_clean[stock_col], terms)
                    col_cond = ~match_mask if is_keep_only else match_mask

                else:
                    apply_map = (ENABLE_FLO_MAPPING and filter_col == 'FLO')
                    vals_to_match = parse_filter_value(
                        clean_val, ORDER_MAPPING.get(filter_col), tolerance_level=0,
                        apply_flo_mapping=apply_map
                    )
                    if vals_to_match:
                        series_to_check = get_normalized_series(stock_col, apply_map=apply_map)
                        clean_vals = [re.sub(r'\s+', '', str(v).upper()) for v in vals_to_match]
                        match_mask = series_to_check.isin(clean_vals)
                        col_cond = ~match_mask if is_keep_only else match_mask
                    else:
                        col_cond = pd.Series(False, index=df_clean.index)

                rule_mask &= col_cond

            except Exception as exc:
                print(f"    WARNING: NO_NEED column '{filter_col}' could not be applied: {exc}")
                # Fail safe: ignore only the problematic condition rather than corrupting the whole run.
                continue

        if has_criteria:
            rows_to_drop_mask |= rule_mask

    # Pair synchronization: if one paired stone fails, remove its mate too.
    def clean_str(val):
        if pd.isna(val) or val == '':
            return None
        s = str(val).strip()
        if s.endswith('.0'):
            s = s[:-2]
        elif s.endswith('.00'):
            s = s[:-3]
        return s

    if 'Extracted_Pair_ID' not in df_clean.columns:
        if 'SGPComment' in df_clean.columns:
            extracted = df_clean['SGPComment'].fillna('').astype(str).str.extract(
                r'(?:PAIRED\s+WITH|MATCH\s+PAIR)\s*[^\d]*([\d\.]+)',
                flags=re.IGNORECASE,
                expand=False
            )
            df_clean['Extracted_Pair_ID'] = extracted.map(clean_str)
        else:
            df_clean['Extracted_Pair_ID'] = None

    if 'Serial No' in df_clean.columns:
        df_clean['Safe_Serial'] = df_clean['Serial No'].map(clean_str)
    else:
        df_clean['Safe_Serial'] = None
    df_clean['Pair_Key'] = df_clean['Extracted_Pair_ID'].map(clean_str)

    failed_serials = set(df_clean.loc[rows_to_drop_mask, 'Safe_Serial'].dropna())
    failed_pair_keys = set(df_clean.loc[rows_to_drop_mask, 'Pair_Key'].dropna())
    failed_serials.discard('')
    failed_pair_keys.discard('')

    additional_fails = df_clean['Pair_Key'].isin(failed_serials) | df_clean['Safe_Serial'].isin(failed_pair_keys)
    keep_mask = ~(rows_to_drop_mask | additional_fails)

    result = df_clean.loc[keep_mask].copy()
    result.drop(columns=[c for c in ['Safe_Serial', 'Pair_Key', 'Extracted_Pair_ID'] if c in result.columns], inplace=True)
    return result

def generate_salesperson_reports(file_path_stock, output_zip_path, filter_file_path, tolerance_level, clean_file_path=None):
    stocks_df = pd.read_excel(file_path_stock, sheet_name='Sheet1', keep_default_na=False)

    if 'Serial No' in stocks_df.columns: stocks_df.drop_duplicates(subset=['Serial No'], keep='first', inplace=True)
    reverse_stock_map = {}
    for final_name, variations in STOCK_COLUMN_MAPPING.items():
        for var in variations: reverse_stock_map[var.upper()] = final_name
    stocks_df.columns = [reverse_stock_map.get(col.strip().upper(), col.strip()) for col in stocks_df.columns]

    # Build expensive normalized/numeric stock Series once and reuse for every party rule.
    stock_filter_cache = build_stock_filter_cache(stocks_df)

    removed_parties, all_change_format_filters = load_filter_data(filter_file_path)

    # --- FILTER FILE CLEANUP: removed parties + duplicate rules only ---
    all_change_format_filters, removed_filter_rows, duplicate_filter_rows = prepare_change_format_filters(
        all_change_format_filters, removed_parties
    )

    # --- NO NEED / CLEAN FILE EXCLUSION LOGIC ---
    exclusion_data = pd.DataFrame()
    if ENABLE_CLEAN_FILE and clean_file_path:
        exclusion_data = load_exclusion_data(clean_file_path)

    global_exclusion_by_sp, specific_exclusion_by_sp_party = build_exclusion_rule_index(exclusion_data)

    salespersons_from_filters = all_change_format_filters['SALES_PERSON'].unique().tolist() if 'SALES_PERSON' in all_change_format_filters.columns else []

    if not salespersons_from_filters:
        print("❌ ERROR: No valid salesperson/filter rows remain after filter cleanup.")
        return

    reports_created = 0

    with zipfile.ZipFile(output_zip_path, 'w', zipfile.ZIP_STORED) as zipf:
        for salesperson in salespersons_from_filters:
            if pd.isna(salesperson): continue
            salesperson_party_reports = {}

            # Indexed NO_NEED lookup: same rules, no repeated full-file scanning.
            global_exclusion_rules = global_exclusion_by_sp.get(salesperson, pd.DataFrame())

            if 'NORMALIZED PARTY' in all_change_format_filters.columns:
                filter_parties = all_change_format_filters[(all_change_format_filters['SALES_PERSON'] == salesperson)]
                for index, party_rule in filter_parties.iterrows():
                    result, party_display_name, country, city = process_party_data(stocks_df, party_rule, tolerance_level, stock_cache=stock_filter_cache)
                    if not result.empty:
                        norm_party = party_rule['NORMALIZED PARTY']
                        if norm_party not in salesperson_party_reports:
                            salesperson_party_reports[norm_party] = {'dataframes': [], 'display_name': party_display_name, 'country': country, 'city': city}
                        salesperson_party_reports[norm_party]['dataframes'].append(result)

            if not salesperson_party_reports: continue

            with BytesIO() as buffer:
                wb = Workbook()
                wb.remove(wb.active)
                sheets_for_nav = []

                for norm_party, party_info in salesperson_party_reports.items():
                    all_results_for_party = pd.concat(party_info['dataframes'], ignore_index=True)
                    if 'Match_Type' in all_results_for_party.columns:
                        all_results_for_party['Match_Type'] = pd.Categorical(all_results_for_party['Match_Type'], categories=['Exact', 'Nearby', 'Maybe'], ordered=True)
                        all_results_for_party.sort_values(by=['Match_Type', 'Shp', 'Carat', 'Color', 'Qua'], inplace=True)

                    all_results_for_party.drop_duplicates(subset=['Serial No'], keep='first', inplace=True)
                    all_results_for_party.reset_index(drop=True, inplace=True)

                    # --- NO NEED / CLEAN FILE EXCLUSION LOGIC ---
                    # Apply after combining all rules for the party so pair exclusion also works
                    # when the two paired stones came from different filter rows.
                    if ENABLE_CLEAN_FILE and not all_results_for_party.empty:
                        before_no_need = len(all_results_for_party)

                        specific_rules = specific_exclusion_by_sp_party.get((salesperson, norm_party), pd.DataFrame())
                        if not global_exclusion_rules.empty and not specific_rules.empty:
                            combined_no_need_rules = pd.concat([global_exclusion_rules, specific_rules], ignore_index=True)
                        elif not global_exclusion_rules.empty:
                            combined_no_need_rules = global_exclusion_rules
                        else:
                            combined_no_need_rules = specific_rules

                        if not combined_no_need_rules.empty:
                            all_results_for_party = apply_exclusion_rules(all_results_for_party, combined_no_need_rules)

                        all_results_for_party.reset_index(drop=True, inplace=True)

                    # --- OUTPUT FILTER CONTROL NEW ---
                    # Applied only after all original matching/sorting/deduplication is complete.
                    # MODE="NONE" returns the original/base output unchanged.
                    party_display_name = party_info['display_name']
                    before_output_filter = len(all_results_for_party)
                    all_results_for_party = apply_match_status_output_filter(all_results_for_party, party_display_name)
                    all_results_for_party.reset_index(drop=True, inplace=True)

                    if not all_results_for_party.empty:
                        if 'Sr.No' in all_results_for_party.columns: all_results_for_party.drop(columns=['Sr.No'], inplace=True)
                        all_results_for_party.insert(0, 'Sr.No', range(1, 1 + len(all_results_for_party)))

                        valid_sheet_name = re.sub(r'[\\/*?:"<>|]', '_', party_display_name)[:31].strip() or "PartySheet"

                        orig_name = valid_sheet_name
                        cnt = 1
                        while valid_sheet_name in [s[0] for s in sheets_for_nav]:
                            valid_sheet_name = f"{orig_name[:28]}_{cnt}"
                            cnt += 1

                        ws = wb.create_sheet(title=valid_sheet_name)
                        df_output = all_results_for_party.copy()

                        # --------------------------------------------------
                        # FIX: NATIVE HYPERLINK LOGIC APPLIED HERE
                        # --------------------------------------------------

                        # 1. Write headers and rows using append/itertuples (much faster than iterrows).
                        ws.append(df_output.columns.tolist())
                        for row in df_output.itertuples(index=False, name=None):
                            ws.append(row)

                        # 3. Apply the native openpyxl hyperlink formatting
                        format_native_hyperlinks(ws, df_output)

                        # --------------------------------------------------

                        apply_styles_to_worksheet(ws, df_output)
                        auto_adjust_column_widths(ws, dataframe=df_output)

                        for h_col in ['Match_Type', 'Is_Filter_Match']:
                            headers = [c.value for c in ws[1]]
                            if h_col in headers: ws.delete_cols(headers.index(h_col) + 1)

                        sheets_for_nav.append((valid_sheet_name, party_display_name, len(df_output), party_info['country'], party_info['city']))

                if sheets_for_nav:
                    nav_ws = wb.create_sheet(title="Navigation", index=0)
                    for c, h in enumerate(["Go To Sheet", "Party Name", "Count", "Country", "City", "Go →"], 1): nav_ws.cell(row=1, column=c, value=h)
                    for r, (sname, pname, count, ctry, city) in enumerate(sheets_for_nav, 2):
                        nav_ws.cell(row=r, column=1, value=sname)
                        nav_ws.cell(row=r, column=2, value=pname)
                        nav_ws.cell(row=r, column=3, value=count)
                        nav_ws.cell(row=r, column=4, value=ctry)
                        nav_ws.cell(row=r, column=5, value=city)
                        nav_ws.cell(row=r, column=6, value="Go →").hyperlink = f"#'{sname}'!A1"
                    apply_styles_to_worksheet(nav_ws)
                    auto_adjust_column_widths(nav_ws)

                # The new output filter may remove every row for this salesperson.
                # In that case, skip creating an empty/invalid workbook.
                if not sheets_for_nav:
                    continue

                wb.save(buffer)
                clean_sp = "".join(x for x in str(salesperson) if x.isalnum() or x in (" ", "_")).strip()[:50] or "Report"
                fname = f"{clean_sp}.xlsx"
                zipf.writestr(fname, buffer.getvalue())
                reports_created += 1

    print("✅ PROCESS COMPLETE")
    print(f"📦 Output file: {output_zip_path}")
    print(f"Removed-party filter rows skipped: {removed_filter_rows:,}")
    print(f"Duplicate filter rows skipped: {duplicate_filter_rows:,}")
    print(f"Salesperson reports created: {reports_created:,}")

if __name__ == "__main__":
    file_path_stock = "/content/drive/MyDrive/Filters File base Only Suggested Stones/Stock Actual.xlsx"
    filter_file_path = "/content/drive/MyDrive/Filters File base Only Suggested Stones/filters.xlsx"
    clean_file_path = "/content/drive/MyDrive/Filters File base Only Suggested Stones/no_need.xlsx"
    output_zip_path = "/content/drive/MyDrive/Filters File base Only Suggested Stones/Regular_Stock_60D_base_no_act&curr_mnth_act.zip"
    try:
        generate_salesperson_reports(
            file_path_stock, output_zip_path, filter_file_path, TOLERANCE_LEVEL, clean_file_path
        )
    except Exception as exc:
        print(f"❌ PROCESS FAILED: {type(exc).__name__}: {exc}")
        raise