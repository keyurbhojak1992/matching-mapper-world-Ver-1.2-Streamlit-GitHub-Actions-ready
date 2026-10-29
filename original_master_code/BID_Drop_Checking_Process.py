import pandas as pd
import numpy as np
from io import StringIO
import re
import os
from original_master_code.inclusion_rules import INCLUSION_ORDER, inclusion_value_matches

# =============================================================================
# 0. CONFIGURATION & GRADE ORDERS
# =============================================================================

COLOR_ORDER = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O-P', 'Q-R', 'S-T', 'U-V', 'W-X', 'Y-Z']

CLARITY_ORDER = ['FL', 'IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2', 'SI3', 'I1', 'I2', 'I3', 'I4', 'I5']

FLO_ORDER = ['NON', 'FNT', 'MED', 'STG', 'VST']

SHADE_ORDER = [
    'NONE',
    'INSPECTION ADVISABLE',
    'MIX TINGE 1',
    'MIX TINGE 2',
    'MIX TINGE 3',
    'LIGHT BROWN',
    'BROWN',
    'STRONG BROWN',
    'VERY STRONG BROWN'
]

MILKY_ORDER = [
    'NONE',
    'LIGHT MILKY',
    'LIGHT CLOUD',
    'MILKY',
    'STRONG MILKY',
    'VERY STRONG MILKY',
    'INSPECTION ADVISABLE'
]

CARAT_ORDER = [
    "0.30-0.39", "0.40-0.49", "0.50-0.69", "0.70-0.89", "0.90-0.99",
    "1.00-1.19", "1.20-1.49", "1.50-1.69", "1.70-1.99", "2.00-2.49",
    "2.50-2.99", "3.00-3.99", "4.00-4.99", "5.00-7.99", "8.00-14.99", "15+"
]

# --- HK Carat Ranges & Order ---
HK_CHINA_RANGES = [
    (0.30, 0.31, "0.30-0.31"), (0.32, 0.39, "0.32-0.39"),
    (0.40, 0.41, "0.40-0.41"), (0.42, 0.49, "0.42-0.49"),
    (0.50, 0.51, "0.50-0.51"), (0.52, 0.59, "0.52-0.59"),
    (0.60, 0.61, "0.60-0.61"), (0.62, 0.69, "0.62-0.69"),
    (0.70, 0.71, "0.70-0.71"), (0.72, 0.79, "0.72-0.79"),
    (0.80, 0.81, "0.80-0.81"), (0.82, 0.89, "0.82-0.89"),
    (0.90, 0.91, "0.90-0.91"), (0.92, 0.99, "0.92-0.99"),
    (1.00, 1.02, "1.00-1.02"), (1.03, 1.19, "1.03-1.19"),
    (1.20, 1.49, "1.20-1.49"), (1.50, 1.59, "1.50-1.59"),
    (1.60, 1.69, "1.60-1.69"), (1.70, 1.79, "1.70-1.79"),
    (1.80, 1.99, "1.80-1.99"), (2.00, 2.10, "2.00-2.10"),
    (2.11, 2.49, "2.11-2.49"), (2.50, 2.59, "2.50-2.59"),
    (2.60, 2.99, "2.60-2.99"), (3.00, 3.10, "3.00-3.10"),
    (3.11, 3.49, "3.11-3.49"), (3.50, 3.59, "3.50-3.59"),
    (3.60, 3.99, "3.60-3.99"), (4.00, 4.99, "4.00-4.99"),
    (5.00, 99.99, "5.00-99.99")
]

HK_CARAT_ORDER = [lbl for low, high, lbl in HK_CHINA_RANGES]

EXCLUDED_SALESPERSONS = [
    'ADMIN',
    'MR. PRIYANK KOTHARI',
    'MR. DHRUVIN HIRPARA',
    'MR. ARUN DHOLIYA',
    'MR. NIKUNJ VASTRAPARA',
    'MR. PRIYAL SHANKAR',
    'MR. SMIT DAYANI',
    'MR. TEJ SHANKAR',
    'MS. LAXMI RATHOD',
    'MR. JATIN MOTISARIYA',
    'MR. RUSHI SHANKAR',
    'MR. AAKASH JAIN',
    'MR. HRITHIK HUNDIA'
]


# =============================================================================
# 0A. HELPER FUNCTIONS
# =============================================================================

def clean_cell_value(x):
    """
    Clean a single Excel cell value.

    Important:
    - Blank / NaN / NULL should become "".
    - NONE should NOT become blank because NONE is a valid diamond value
      for Shade, Milky, TBINC_NAME, SBINC_NAME, TOIN_NAME, etc.
    """
    if pd.isna(x):
        return ""

    s = str(x).strip().upper()
    s = re.sub(r'\s+', ' ', s)

    if s in ["", "NAN", "NULL", "NAT"]:
        return ""

    return s


def normalize_party_name(name):
    """
    Normalize party names so AGRAWAL DIAMOND and Agrawal Diamond match.
    """
    if not isinstance(name, str) or pd.isna(name):
        return ""

    name = re.sub(r'[^\w\s]', '', name).upper().strip()
    name = re.sub(r'\b(INC|LLC|LTD|CO|CORP|COMPANY)\b', '', name).strip()
    name = re.sub(r'\s+', ' ', name)

    return name


def get_neighbors(val, order_list, tolerance_level=1):
    """
    Get neighboring grades based on tolerance.
    Example: Color G with tolerance 1 -> F, G, H.
    """
    val = str(val).strip().upper()

    if val not in order_list or tolerance_level == 0:
        return [val]

    idx = order_list.index(val)
    start = max(0, idx - tolerance_level)
    end = min(len(order_list) - 1, idx + tolerance_level)

    return order_list[start:end + 1]


def get_hk_carat_range(carat_val):
    """
    Special HK/China carat grouping.
    """
    if pd.isna(carat_val) or carat_val == '':
        return "UNKNOWN"

    try:
        val = float(carat_val)
    except Exception:
        return "UNKNOWN"

    for low, high, label in HK_CHINA_RANGES:
        if low <= val <= high:
            return label

    return "5.00-99.99" if val >= 5.00 else "UNKNOWN"


def standardize_rule_col_name(col):
    """
    Standardizes columns from no_need.xlsx, filters.xlsx, tolerance sheet, removed party sheet.

    Important:
    - TBINC NAME, TBINC_NAME, TB INC NAME -> TBINC_NAME
    - SBINC NAME, SBINC_NAME, SB INC NAME -> SBINC_NAME
    - Ratio -> RATIO
    - Color Name -> COLOR NAME
    """
    original = str(col).strip().upper()

    c = original
    c = re.sub(r'\s+', '_', c)
    c = c.replace('-', '_')
    c = re.sub(r'_+', '_', c)

    # Remove dots only for alias lookup.
    c_no_dot = c.replace('.', '')

    aliases = {
        'SOLD_PARTY': 'SOLD PARTY',
        'PARTY': 'SOLD PARTY',
        'PARTY_NAME': 'SOLD PARTY',

        'SALES_PERSON': 'SALES_PERSON',
        'SALE_PERSON': 'SALES_PERSON',
        'SALESPERSON': 'SALES_PERSON',

        'COLOR_NAME': 'COLOR NAME',
        'COLOR': 'COLOR NAME',

        'QUA': 'QUA',
        'QUALITY': 'QUA',
        'CLARITY': 'QUA',

        'SHP': 'SHP',
        'SHAPE': 'SHP',

        'CARAT': 'CARAT',
        'RATE': 'RATE',
        'DISC%': 'DISC%',
        'DISCOUNT': 'DISC%',
        'RATIO': 'RATIO',
        'TAB': 'TAB',
        'TABLE': 'TAB',
        'CA': 'C/A',
        'C_A': 'C/A',

        'CUT': 'CUT',
        'POL': 'POL',
        'SYM': 'SYM',
        'FLO': 'FLO',
        'FLUO': 'FLO',
        'FLUORESCENCE': 'FLO',
        'LAB': 'LAB',

        'SHADE': 'SHADE',
        'MILKY': 'MILKY',

        'EFACET': 'E.FACET',
        'E_FACET': 'E.FACET',
        'EYE_CLEAN': 'EYE CLEAN',
        'EYECLEAN': 'EYE CLEAN',
        'GRAINING': 'GRAINING',
        'CULET': 'CULET',
        'LUSTER': 'LUSTER',

        'TINC_NAME': 'TINC_NAME',
        'TINCNAME': 'TINC_NAME',
        'T_INC_NAME': 'TINC_NAME',

        'SINC_NAME': 'SINC_NAME',
        'SINCNAME': 'SINC_NAME',
        'S_INC_NAME': 'SINC_NAME',

        'TOIN_NAME': 'TOIN_NAME',
        'TOINNAME': 'TOIN_NAME',
        'TO_INC_NAME': 'TOIN_NAME',

        'SOIN_NAME': 'SOIN_NAME',
        'SOINNAME': 'SOIN_NAME',
        'SO_INC_NAME': 'SOIN_NAME',

        'GIRDLE_OPEN_NAME': 'GIRDLE_OPEN_NAME',
        'GIRDLEOPENNAME': 'GIRDLE_OPEN_NAME',

        'PAV_OPEN_NAME': 'PAV_OPEN_NAME',
        'PAVOPENNAME': 'PAV_OPEN_NAME',

        'TBINC_NAME': 'TBINC_NAME',
        'TBINCNAME': 'TBINC_NAME',
        'TB_INC_NAME': 'TBINC_NAME',
        'TB_INC': 'TBINC_NAME',
        'TABLE_BLACK': 'TBINC_NAME',

        'SBINC_NAME': 'SBINC_NAME',
        'SBINCNAME': 'SBINC_NAME',
        'SB_INC_NAME': 'SBINC_NAME',
        'SB_INC': 'SBINC_NAME',
        'SIDE_BLACK': 'SBINC_NAME',

        'TABLE_INCL': 'TABLE INCL',
        'MEASUREMENT': 'MEASUREMENT',
        'SGP_COMMENT': 'SGP COMMENT',
        'KEYTOSYMBOL': 'KEYTOSYMBOL',
        'KEY_TO_SYMBOL': 'KEYTOSYMBOL',
        'COMMENT': 'COMMENT',
        'COLOR_DESC': 'COLOR DESC',
        'REMARKS': 'REMARKS',

        'TOLERANCE_+_2': 'TOLERANCE +-2',
        'TOLERANCE_+-2': 'TOLERANCE +-2',
        'TOLERANCE_+_1': 'TOLERANCE +-1',
        'TOLERANCE_+-1': 'TOLERANCE +-1',
        'CARAT_GRP_TOLERANCE': 'CARAT_GRP_TOLERANCE'
    }

    if c in aliases:
        return aliases[c]

    if c_no_dot in aliases:
        return aliases[c_no_dot]

    return original


def standardize_rule_columns(df):
    """
    Apply rule-column standardization to no_need, filters, tolerance, removed party.
    """
    if df is None or df.empty:
        return pd.DataFrame()

    df = df.copy()
    df.columns = [standardize_rule_col_name(c) for c in df.columns]

    return df


def clean_dataframe(df):
    """
    Clean main dataframes: bids, sales, suggested, stock.
    Keeps normal business column names mostly unchanged.
    """
    df = df.loc[:, ~df.columns.duplicated()].copy()

    if 'Serial No' in df.columns:
        df['Serial No'] = df['Serial No'].apply(
            lambda x: str(x).replace('.0', '').strip().upper() if pd.notnull(x) else str(x)
        )

    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].astype(str).str.strip().str.upper()

    return df


def add_stock_alias_columns(df_stock):
    """
    Ensures stock columns exist with the exact names used by matching logic.
    Handles Excel variations like TBINC NAME vs TBINC_NAME.
    """
    if df_stock is None or df_stock.empty:
        return pd.DataFrame()

    df_stock = df_stock.copy()

    alias_map = {}

    for col in df_stock.columns:
        std = standardize_rule_col_name(col)

        if std == 'TBINC_NAME':
            alias_map[col] = 'TBINC_NAME'
        elif std == 'SBINC_NAME':
            alias_map[col] = 'SBINC_NAME'
        elif std == 'TOIN_NAME':
            alias_map[col] = 'TOIN_NAME'
        elif std == 'SOIN_NAME':
            alias_map[col] = 'SOIN_NAME'
        elif std == 'GIRDLE_OPEN_NAME':
            alias_map[col] = 'GIRDLE_OPEN_NAME'
        elif std == 'PAV_OPEN_NAME':
            alias_map[col] = 'PAV_OPEN_NAME'
        elif std == 'E.FACET':
            alias_map[col] = 'E.Facet'
        elif std == 'EYE CLEAN':
            alias_map[col] = 'EYE Clean'
        elif std == 'RATIO':
            alias_map[col] = 'Ratio'
        elif std == 'TAB':
            alias_map[col] = 'Tab'
        elif std == 'C/A':
            alias_map[col] = 'C/A'

    for old_col, new_col in alias_map.items():
        if new_col not in df_stock.columns:
            df_stock[new_col] = df_stock[old_col]

    return df_stock


def check_numeric_scalar(val, rule_str):
    """
    Numeric rule checker.
    Supports:
    - < 1.90
    - <= 1.90
    - > 1.90
    - >= 1.90
    - != 1.90
    - 1.90
    - 1.50 TO 1.90
    - multiple comma rules
    """
    if pd.isna(rule_str) or str(rule_str).strip().upper() in ['', 'NAN', 'NONE']:
        return True

    try:
        v = float(val)
    except Exception:
        return False

    r = str(rule_str).strip().upper()

    for sr in [x.strip() for x in r.split(',') if x.strip()]:
        try:
            if 'TO' in sr:
                p = sr.split('TO')
                low = float(p[0].strip())
                high = float(p[1].strip())

                if low <= v <= high:
                    return True

            elif sr.startswith('>='):
                if v >= float(sr[2:].strip()):
                    return True

            elif sr.startswith('>'):
                if v > float(sr[1:].strip()):
                    return True

            elif sr.startswith('<='):
                if v <= float(sr[2:].strip()):
                    return True

            elif sr.startswith('<'):
                if v < float(sr[1:].strip()):
                    return True

            elif sr.startswith('!='):
                if v != float(sr[2:].strip()):
                    return True

            else:
                if v == float(sr.strip()):
                    return True

        except Exception:
            continue

    return False


def expand_string_rule(rule_str, order_list=None):
    """
    Expands text rules.
    Supports:
    - G TO I
    - SI1 TO I1
    - MINOR BLK, MEDIUM BLK
    """
    rule_str = clean_cell_value(rule_str)

    if not rule_str:
        return []

    values = []

    if order_list and 'TO' in rule_str:
        parts = rule_str.split('TO')

        if len(parts) == 2:
            s_val = parts[0].strip()
            e_val = parts[1].strip()

            if s_val in order_list and e_val in order_list:
                s_idx = order_list.index(s_val)
                e_idx = order_list.index(e_val)

                if s_idx <= e_idx:
                    values.extend(order_list[s_idx:e_idx + 1])
                else:
                    values.extend(order_list[e_idx:s_idx + 1])
            else:
                values.append(rule_str)
        else:
            values.append(rule_str)

    elif ',' in rule_str:
        values.extend([clean_cell_value(x) for x in rule_str.split(',') if clean_cell_value(x)])

    else:
        values.append(rule_str)

    return list(dict.fromkeys(values))


def split_rule_values(rule_val):
    """
    Split comma-based no_need values.
    """
    rule_val = clean_cell_value(rule_val)

    if not rule_val:
        return []

    return [clean_cell_value(x) for x in rule_val.split(',') if clean_cell_value(x)]


# =============================================================================
# 1. LOAD & CLEAN DATA
# =============================================================================



def run_bid_drop_check(
    *,
    suggested_path,
    bids_path,
    sales_path,
    stock_path,
    filters_path,
    no_need_path,
    output_path,
):
    """Run the BID drop diagnostic using explicit application paths."""
    print("Loading core data...")

    df_suggested = pd.read_excel(suggested_path)
    df_bids = pd.read_excel(bids_path)
    df_sales = pd.read_excel(sales_path)
    df_stock = pd.read_excel(stock_path)

    print("Loading filters.xlsx...")

    try:
        filters_path = str(filters_path)

        xl_filters = pd.ExcelFile(filters_path)
        available_sheets = xl_filters.sheet_names

    except Exception as e:
        print("filters.xlsx not found or not readable:", e)
        xl_filters = None
        available_sheets = []


    def load_sheet_safe(xl_file, expected_name, available):
        """
        Load a sheet by case-insensitive name.
        """
        if xl_file is None:
            return pd.DataFrame()

        for actual_name in available:
            if actual_name.strip().upper() == expected_name.strip().upper():
                df = xl_file.parse(actual_name)
                return standardize_rule_columns(df)

        return pd.DataFrame()


    if available_sheets:
        df_removed_parties = load_sheet_safe(xl_filters, 'Removed Party', available_sheets)
        df_tolerance = load_sheet_safe(xl_filters, 'Tolerance', available_sheets)
        df_filters = load_sheet_safe(xl_filters, 'Change Format', available_sheets)
    else:
        df_removed_parties = pd.DataFrame()
        df_tolerance = pd.DataFrame()
        df_filters = pd.DataFrame()


    print("Loading no_need.xlsx...")

    try:
        no_need_path = str(no_need_path)

        df_no_need = pd.read_excel(no_need_path)
        df_no_need = standardize_rule_columns(df_no_need)

    except Exception as e:
        print("no_need.xlsx not found or not readable:", e)
        df_no_need = pd.DataFrame()


    # --- Rename main data columns ---
    df_bids.rename(columns={
        'Qua': 'Clarity',
        'Party Name': 'Party'
    }, inplace=True)

    df_sales.rename(columns={
        'Col': 'Color',
        'Qua': 'Clarity',
        'Sold Party': 'Party'
    }, inplace=True)

    df_stock.rename(columns={
        'Qua': 'Clarity'
    }, inplace=True)

    df_suggested.rename(columns={
        'Sold Party': 'Party'
    }, inplace=True)


    # --- Clean main dataframes ---
    df_bids = clean_dataframe(df_bids)
    df_sales = clean_dataframe(df_sales)
    df_suggested = clean_dataframe(df_suggested)
    df_stock = clean_dataframe(df_stock)
    df_stock = add_stock_alias_columns(df_stock)


    # --- Clean rule dataframe values ---
    for df_rule in [df_removed_parties, df_tolerance, df_filters, df_no_need]:
        if df_rule is not None and not df_rule.empty:
            for col in df_rule.columns:
                if df_rule[col].dtype == 'object':
                    df_rule[col] = df_rule[col].astype(str).str.strip().str.upper()


    # =============================================================================
    # 1A. MAP SALESPERSONS AND COUNTRIES
    # =============================================================================

    sales_sp_col = next(
        (c for c in df_sales.columns if c.upper() in ['SALE_PERSON', 'SALES_PERSON', 'SALES PERSON']),
        None
    )

    if sales_sp_col:
        party_to_salesperson = (
            df_sales
            .groupby(df_sales['Party'].apply(normalize_party_name))[sales_sp_col]
            .first()
            .astype(str)
            .str.strip()
            .str.upper()
            .to_dict()
        )
    else:
        party_to_salesperson = {}


    sales_country_col = next(
        (c for c in df_sales.columns if c.upper() == 'COUNTRY'),
        None
    )

    if sales_country_col:
        party_to_country = (
            df_sales
            .groupby(df_sales['Party'].apply(normalize_party_name))[sales_country_col]
            .first()
            .astype(str)
            .str.strip()
            .str.upper()
            .to_dict()
        )
    else:
        party_to_country = {}


    # --- Removed party dict ---
    removed_parties_dict = {}

    if not df_removed_parties.empty and 'SOLD PARTY' in df_removed_parties.columns:
        for _, row in df_removed_parties.iterrows():
            party_name = str(row['SOLD PARTY'])

            if party_name.upper() not in ['NAN', 'NAT', 'NONE', '']:
                norm_name = normalize_party_name(party_name)
                remark = str(row.get('REMARKS', '')).strip()

                if remark.upper() in ['NAN', 'NAT', 'NONE', '']:
                    remark = "Listed in sheet (No Remarks)"

                removed_parties_dict[norm_name] = remark


    # --- Add normalized party columns in rule sheets ---
    if not df_tolerance.empty and 'SOLD PARTY' in df_tolerance.columns:
        df_tolerance['NORMALIZED PARTY'] = df_tolerance['SOLD PARTY'].astype(str).apply(normalize_party_name)

    if not df_filters.empty and 'SOLD PARTY' in df_filters.columns:
        df_filters['NORMALIZED PARTY'] = df_filters['SOLD PARTY'].astype(str).apply(normalize_party_name)

    if not df_no_need.empty and 'SOLD PARTY' in df_no_need.columns:
        df_no_need['NORMALIZED PARTY'] = df_no_need['SOLD PARTY'].astype(str).apply(normalize_party_name)


    # --- Suggested pairs ---
    suggested_pairs = set(
        zip(
            df_suggested['Serial No'].astype(str),
            df_suggested['Party'].apply(normalize_party_name)
        )
    )

    stock_serials = set(df_stock['Serial No'].astype(str))

    df_stock_unique = df_stock.drop_duplicates(subset=['Serial No'], keep='first')
    stock_details_map = df_stock_unique.set_index('Serial No').to_dict('index')


    def make_supertag(row):
        return (
            f"{row.get('Shp', '')}-"
            f"{row.get('Carat_grp', '')}-"
            f"{row.get('Color', '')}-"
            f"{row.get('Clarity', '')}-"
            f"{row.get('Flo', '')}-"
            f"{row.get('Lab', '')}"
        )


    df_sales['Past_Supertag_6'] = df_sales.apply(make_supertag, axis=1)
    df_sales['Norm_Party'] = df_sales['Party'].apply(normalize_party_name)

    past_sales_by_party = df_sales.groupby('Norm_Party')


    # =============================================================================
    # 2. FILTER / NO_NEED EVALUATION ENGINE
    # =============================================================================

    def evaluate_no_need_row(stone, rule):
        """
        Correct no_need evaluator.

        Rule behavior:
        - Every non-empty cell in the no_need row must match.
        - Example:
            Party = SHIVAANSH JEWELS
            SHP = MQ
            RATIO = < 1.90
          Means block only when Shp is MQ AND Ratio < 1.90.

        - Example:
            Party = AGRAWAL DIAMOND
            TBINC_NAME = MINOR BLK, NOTICEABLE BLACK, MEDIUM BLK
          Means block if TBINC_NAME is one of those values.
        """
        matched_reasons = []
        has_any_condition = False

        col_map = {
            'SHP': ('Shp', 'text', None),
            'CARAT': ('Carat', 'numeric', None),
            'RATE': ('Rate', 'numeric', None),
            'DISC%': ('Disc%', 'numeric', None),
            'RATIO': ('Ratio', 'numeric', None),
            'TAB': ('Tab', 'numeric', None),
            'C/A': ('C/A', 'numeric', None),

            'COLOR NAME': ('Color', 'text', COLOR_ORDER),
            'QUA': ('Clarity', 'text', CLARITY_ORDER),
            'CUT': ('Cut', 'text', ['EX', 'VG', 'G', 'F']),
            'POL': ('Pol', 'text', ['EX', 'VG', 'G', 'F']),
            'SYM': ('Sym', 'text', ['EX', 'VG', 'G', 'F']),
            'FLO': ('Flo', 'text', FLO_ORDER),
            'LAB': ('Lab', 'text', None),
            'SHADE': ('Shade', 'text', SHADE_ORDER),
            'MILKY': ('Milky', 'text', MILKY_ORDER),

            'E.FACET': ('E.Facet', 'text', None),
            'CULET': ('Culet', 'text', None),
            'LUSTER': ('Luster', 'text', None),
            'EYE CLEAN': ('EYE Clean', 'text', None),
            'GRAINING': ('Graining', 'text', None),

            'TINC_NAME': ('TINC_NAME', 'text', None),
            'SINC_NAME': ('SINC_NAME', 'text', None),
            'TOIN_NAME': ('TOIN_NAME', 'text', None),
            'SOIN_NAME': ('SOIN_NAME', 'text', None),
            'GIRDLE_OPEN_NAME': ('GIRDLE_OPEN_NAME', 'text', None),
            'PAV_OPEN_NAME': ('PAV_OPEN_NAME', 'text', None),
            'TBINC_NAME': ('TBINC_NAME', 'text', None),
            'SBINC_NAME': ('SBINC_NAME', 'text', None),
        }

        for rule_col, (stock_col, col_type, order) in col_map.items():
            rule_val = clean_cell_value(rule.get(rule_col, ""))

            if not rule_val:
                continue

            has_any_condition = True

            if col_type == 'numeric':
                if not check_numeric_scalar(stone.get(stock_col, ""), rule_val):
                    return False, []

                matched_reasons.append(f"{rule_col} {rule_val}")
                continue

            stone_val = clean_cell_value(stone.get(stock_col, ""))

            # Support ! keep-only logic:
            # Example: !NONE means block if stone value is NOT NONE.
            is_keep_only = rule_val.startswith('!')
            clean_rule = rule_val[1:].strip() if is_keep_only else rule_val

            target_values = expand_string_rule(clean_rule, order)

            if is_keep_only:
                if stone_val not in target_values:
                    matched_reasons.append(
                        f"{rule_col} required {clean_rule}, but stone has {stone_val or 'BLANK'}"
                    )
                else:
                    return False, []
            else:
                if stone_val not in target_values:
                    return False, []

                matched_reasons.append(f"{rule_col} matched blocked value {stone_val}")

        if has_any_condition and matched_reasons:
            return True, matched_reasons

        return False, []


    def evaluate_advanced_row(stone, rule):
        """
        filters.xlsx evaluator.

        Remarks = INCLUDE means stone must match at least one INCLUDE rule.
        Remarks = EXCLUDE means stone is blocked if it matches EXCLUDE rule.

        This now includes RATIO also.
        """
        has_criteria = False
        reasons = []

        col_map = {
            'SHP': ('Shp', 'text', None),
            'QUA': ('Clarity', 'text', CLARITY_ORDER),
            'CUT': ('Cut', 'text', ['EX', 'VG', 'G', 'F']),
            'POL': ('Pol', 'text', ['EX', 'VG', 'G', 'F']),
            'SYM': ('Sym', 'text', ['EX', 'VG', 'G', 'F']),
            'FLO': ('Flo', 'text', FLO_ORDER),
            'LAB': ('Lab', 'text', None),
            'DISC%': ('Disc%', 'numeric', None),
            'SHADE': ('Shade', 'text', SHADE_ORDER),
            'MILKY': ('Milky', 'text', MILKY_ORDER),
            'COLOR NAME': ('Color', 'text', COLOR_ORDER),
            'CARAT': ('Carat', 'numeric', None),
            'RATE': ('Rate', 'numeric', None),
            'RATIO': ('Ratio', 'numeric', None),
            'TAB': ('Tab', 'numeric', None),
            'C/A': ('C/A', 'numeric', None),
            'TINC_NAME': ('TINC_NAME', 'text', INCLUSION_ORDER),
            'SINC_NAME': ('SINC_NAME', 'text', INCLUSION_ORDER),
            'TBINC_NAME': ('TBINC_NAME', 'text', None),
            'SBINC_NAME': ('SBINC_NAME', 'text', None),
            'TOIN_NAME': ('TOIN_NAME', 'text', None),
            'SOIN_NAME': ('SOIN_NAME', 'text', None),
            'GIRDLE_OPEN_NAME': ('GIRDLE_OPEN_NAME', 'text', None),
            'PAV_OPEN_NAME': ('PAV_OPEN_NAME', 'text', None),
        }

        for rule_col, (stock_col, col_type, order) in col_map.items():
            rule_val = clean_cell_value(rule.get(rule_col, ""))

            if not rule_val:
                continue

            has_criteria = True

            if col_type == 'numeric':
                if not check_numeric_scalar(stone.get(stock_col, ""), rule_val):
                    return False, []

                reasons.append(f"{rule_col} {rule_val}")
                continue

            stone_val = clean_cell_value(stone.get(stock_col, ""))

            if rule_col in {'TINC_NAME', 'SINC_NAME'}:
                if not inclusion_value_matches(stone_val, rule_val):
                    return False, []
                reasons.append(f'{rule_col} matches {rule_val}')
                continue

            parts = [p.strip() for p in rule_val.split(',') if p.strip()]
            inc_rules = [p for p in parts if not p.startswith('!')]
            exc_rules = [p.lstrip('!') for p in parts if p.startswith('!')]

            if inc_rules:
                inc_vals = []

                for ir in inc_rules:
                    inc_vals.extend(expand_string_rule(ir, order))

                if stone_val not in inc_vals:
                    return False, []

                reasons.append(f"{rule_col} matches {inc_rules}")

            if exc_rules:
                exc_vals = []

                for er in exc_rules:
                    exc_vals.extend(expand_string_rule(er, order))

                if stone_val in exc_vals:
                    return False, []

                reasons.append(f"{rule_col} passes exclusion {exc_rules}")

        return has_criteria, reasons


    def check_filters_detailed(stone_dict, party_norm_name):
        """
        Filter priority:
        1. no_need.xlsx
        2. filters.xlsx
        """

        # --- 1. no_need.xlsx ---
        if not df_no_need.empty and 'NORMALIZED PARTY' in df_no_need.columns:
            app_no_need = df_no_need[
                (df_no_need['NORMALIZED PARTY'] == party_norm_name) |
                (df_no_need['NORMALIZED PARTY'] == 'ALL')
            ]

            for _, rule in app_no_need.iterrows():
                is_match, reasons = evaluate_no_need_row(stone_dict, rule)

                if is_match:
                    return (
                        True,
                        "NO_NEED.XLSX: " + " AND ".join(reasons),
                        False,
                        ""
                    )

        # --- 2. filters.xlsx ---
        if not df_filters.empty and 'NORMALIZED PARTY' in df_filters.columns:
            party_rules = df_filters[df_filters['NORMALIZED PARTY'] == party_norm_name]

            if not party_rules.empty:
                passed_include = False
                has_include_rules = False

                for _, rule in party_rules.iterrows():
                    remarks = clean_cell_value(rule.get('REMARKS', 'INCLUDE'))

                    is_match, reasons = evaluate_advanced_row(stone_dict, rule)

                    if remarks == 'INCLUDE':
                        has_include_rules = True

                        if is_match:
                            passed_include = True

                    elif remarks == 'EXCLUDE':
                        if is_match:
                            return (
                                False,
                                "",
                                True,
                                "FILTERS.XLSX: Matched EXCLUDE rule -> (" + " AND ".join(reasons) + ")"
                            )

                if has_include_rules and not passed_include:
                    return (
                        False,
                        "",
                        True,
                        "FILTERS.XLSX: Stone failed all INCLUDE bounds set for party."
                    )

        return False, "", False, ""


    # =============================================================================
    # 2A. OPTIONAL DEBUG CHECK
    # =============================================================================

    print("\nRule sheet column check:")
    if not df_no_need.empty:
        print("no_need columns:", df_no_need.columns.tolist())
    else:
        print("no_need.xlsx is empty or not loaded.")

    if not df_filters.empty:
        print("filters columns:", df_filters.columns.tolist())
    else:
        print("filters.xlsx Change Format sheet is empty or not loaded.")


    # =============================================================================
    # 3. RUN THE STRICT DIAGNOSTIC LOOP
    # =============================================================================

    print("\nRunning diagnostic engine...")

    report_data = []

    for idx, bid_row in df_bids.iterrows():
        stone_id = str(bid_row['Serial No'])
        party_raw = bid_row['Party']
        party_norm = normalize_party_name(party_raw)

        sp = party_to_salesperson.get(party_norm, '')
        country = party_to_country.get(party_norm, '')

        is_hk_party = country in ['HONGKONG', 'HK', 'H.K', 'CHINA']

        status = ""
        reason = ""
        stock_info = ""

        # --- Stock details ---
        stock_stone = stock_details_map.get(stone_id, {})

        shp = stock_stone.get('Shp', '')
        col = stock_stone.get('Color', '')
        cla = stock_stone.get('Clarity', '')
        flo = stock_stone.get('Flo', '')
        lab = stock_stone.get('Lab', '')
        exact_carat = stock_stone.get('Carat', 0)
        carat_grp = stock_stone.get('Carat_grp', '')

        # --- HK carat range override ---
        if is_hk_party:
            hk_grp = get_hk_carat_range(exact_carat)

            if hk_grp != "UNKNOWN":
                carat_grp = hk_grp

        carat_order_to_use = HK_CARAT_ORDER if is_hk_party else CARAT_ORDER

        # =============================================================================
        # DECISION TREE
        # =============================================================================

        if (stone_id, party_norm) in suggested_pairs:
            status = "✅ SUCCESS"
            reason = "Stone suggested and converted to bid."

        elif stone_id not in stock_serials:
            status = "⏳ OUT OF STOCK"
            reason = "Stone wasn't in stock.xlsx when algo ran."

        elif party_norm in removed_parties_dict:
            rm_remark = removed_parties_dict[party_norm]

            status = "🚫 BLOCKED: REMOVED PARTY"
            reason = f"Party is in 'Removed Party' list (Remarks: {rm_remark})."

        elif sp in EXCLUDED_SALESPERSONS:
            status = "🚫 BLOCKED: EXC. SALES PERSON"
            reason = f"Exc.Sales Person ({sp.title()})"

        else:
            current_tolerance = 1
            carat_tolerance_enabled = False

            # --- Tolerance rules ---
            if not df_tolerance.empty and 'NORMALIZED PARTY' in df_tolerance.columns:
                tol_rule = df_tolerance[df_tolerance['NORMALIZED PARTY'] == party_norm]

                if not tol_rule.empty:
                    if clean_cell_value(tol_rule.iloc[0].get('TOLERANCE +-2', '')) == 'YES':
                        current_tolerance = 2

                    elif clean_cell_value(tol_rule.iloc[0].get('TOLERANCE +-1', '')) == 'NO':
                        current_tolerance = 0

                    if clean_cell_value(tol_rule.iloc[0].get('CARAT_GRP_TOLERANCE', '')) == 'YES':
                        carat_tolerance_enabled = True

            # --- no_need / filters ---
            fails_no_need, rsn_no_need, fails_adv, rsn_adv = check_filters_detailed(stock_stone, party_norm)

            shade = stock_stone.get('Shade', 'NONE')
            tb_inc = stock_stone.get('TBINC_NAME', 'NONE')
            sb_inc = stock_stone.get('SBINC_NAME', 'NONE')
            ratio = stock_stone.get('Ratio', '')

            carat_clue = "Yes" if carat_tolerance_enabled else "No"

            stock_info = (
                f"Tol: {current_tolerance} | "
                f"Crt-Tol: {carat_clue} | "
                f"Shade: {shade} | "
                f"Ratio: {ratio} | "
                f"TB: {tb_inc} | "
                f"SB: {sb_inc}"
            )

            if fails_no_need:
                status = "🛑 BLOCKED: NO_NEED"
                reason = rsn_no_need

            elif fails_adv:
                status = "🛑 BLOCKED: FILTERS.XLSX"
                reason = rsn_adv

            else:
                party_has_history = party_norm in past_sales_by_party.groups

                if not party_has_history:
                    status = "🚀 NEW DEMAND"
                    reason = "REJECTED: Party has NO past sales history in Sales & Actions."

                else:
                    party_sales = past_sales_by_party.get_group(party_norm)

                    has_tolerant_match = False
                    matched_past_tag = ""

                    valid_carats = (
                        get_neighbors(carat_grp, carat_order_to_use, 1)
                        if carat_tolerance_enabled
                        else [carat_grp]
                    )

                    base_match = party_sales[
                        (party_sales['Shp'] == shp) &
                        (party_sales['Carat_grp'].isin(valid_carats)) &
                        (party_sales['Lab'] == lab)
                    ]

                    valid_cols = get_neighbors(col, COLOR_ORDER, current_tolerance)
                    valid_clas = get_neighbors(cla, CLARITY_ORDER, current_tolerance)
                    valid_flos = get_neighbors(flo, FLO_ORDER, current_tolerance)

                    if not base_match.empty:
                        tol_match = base_match[
                            (base_match['Color'].isin(valid_cols)) &
                            (base_match['Clarity'].isin(valid_clas)) &
                            (base_match['Flo'].isin(valid_flos))
                        ]

                        if not tol_match.empty:
                            has_tolerant_match = True
                            matched_past_tag = tol_match.iloc[0]['Past_Supertag_6']

                    if has_tolerant_match:
                        status = "⚠️ ALGO MISS (OTHER)"
                        reason = (
                            f"Passed filters & matched past sales ({matched_past_tag}), "
                            f"but dropped by Top-N/Capacity limits."
                        )

                    else:
                        status = "🚀 NEW DEMAND"

                        missing_global = []

                        if shp not in party_sales['Shp'].values:
                            missing_global.append(f"Shape({shp})")

                        if not party_sales['Carat_grp'].isin(valid_carats).any():
                            missing_global.append(f"Carat({carat_grp})")

                        if lab not in party_sales['Lab'].values:
                            missing_global.append(f"Lab({lab})")

                        if missing_global:
                            verb = "are" if len(missing_global) > 1 else "is"

                            reason = (
                                f"REJECTED: {', '.join(missing_global)} {verb} "
                                f"completely missing across all their past sales."
                            )

                        else:
                            if base_match.empty:
                                crt_str = (
                                    f"Carat({carat_grp} ±Tol)"
                                    if carat_tolerance_enabled
                                    else f"Carat({carat_grp})"
                                )

                                reason = (
                                    f"REJECTED: Party buys Shape({shp}), {crt_str}, Lab({lab}) "
                                    f"on separate stones, but NEVER combined together."
                                )

                            else:
                                failed_tol = []

                                if not base_match['Color'].isin(valid_cols).any():
                                    failed_tol.append(f"Color({col})")

                                if not base_match['Clarity'].isin(valid_clas).any():
                                    failed_tol.append(f"Clarity({cla})")

                                if not base_match['Flo'].isin(valid_flos).any():
                                    failed_tol.append(f"Flo({flo})")

                                crt_label = (
                                    "Carat Tolerance Match"
                                    if carat_tolerance_enabled and carat_grp not in base_match['Carat_grp'].values
                                    else carat_grp
                                )

                                if failed_tol:
                                    verb = "are" if len(failed_tol) > 1 else "is"

                                    reason = (
                                        f"Base Match: {shp}-[{crt_label}]-{lab}. "
                                        f"REJECTED: {', '.join(failed_tol)} {verb} "
                                        f"completely outside their tolerance (±{current_tolerance})."
                                    )

                                else:
                                    reason = (
                                        f"Base Match: {shp}-[{crt_label}]-{lab}. "
                                        f"REJECTED: Col({col}), Cla({cla}), and Flo({flo}) "
                                        f"pass tolerance individually, but NEVER combined on a single stone."
                                    )

        # --- Dynamic bid tag ---
        bid_tag = f"{shp}-{carat_grp}-{col}-{cla}-{flo}-{lab}"

        report_data.append({
            'Serial No': stone_id,
            'Party Name': party_raw,
            'Bid Supertag': bid_tag,
            'Diagnostic Status': status,
            'Explanation': reason,
            'Stock Clues': stock_info
        })


    # =============================================================================
    # 4. OUTPUT
    # =============================================================================

    df_diagnostic = pd.DataFrame(report_data)

    output_file = str(output_path)
    df_diagnostic.to_excel(output_file, index=False)

    print(f"✅ Output successfully saved as '{output_file}'")

    return output_file