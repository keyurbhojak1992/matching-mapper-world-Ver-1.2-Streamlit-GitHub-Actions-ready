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
from original_master_code.inclusion_rules import INCLUSION_COLUMNS, INCLUSION_ORDER, inclusion_condition

# --- USER INPUT SECTION ---
# ==========================
# Set the tolerance level for matching:
# 1: for a +/-1 tolerance (Recommended)
TOLERANCE_LEVEL = 1

# --- FLUORESCENCE CONTROL ---
# True: Groups VSL, SL, FNT together AND makes NON next to FNT.
#       Request 'NON' -> Matches NON + FNT (VSL/SL).
# False: Strict distinct order (NON -> VSL -> SL -> FNT).
ENABLE_FLO_MAPPING = False

# --- RELATED TAG CONTROL ---
# True: Adds the 'Tag_No' column to the output file showing the tag from the filters file.
# False: Standard behavior, no new column added.
ENABLE_RELATED_TAG = True
# ==========================

COLOR_ORDER = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O-P', 'Q-R', 'S-T', 'U-V', 'W-X', 'Y-Z']
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
TOIN_SOIN_GIRDLE_PAV_OPEN_ORDER = ['NONE', 'HAIR LINE', 'SMALL OPEN', 'MEDIUM OPEN', 'HEAVY OPEN']
TBINC_SBINC_ORDER = ['NONE', 'PIN POINT BLK', 'MINOR BLK', 'NOTICABLE BLACK', 'MEDIUM BLK', 'HEAVY BLK']

# New order list for sorting
SHAPE_ORDER = [
    "RBC", "PEAR", "PMB", "OVAL", "OMB", "MQ", "MMB", "EME", "SEM", "RMB", "RM", "ECB", "ECMB", "ECM", "ECBF", "CB", "CMB", "CM", "HRT", "HMB", "PR", "TR", "BGT", "PMC", "OMC", "MMC", "SRMB", "SRM", "TP", "SLD", "CAD", "SCAD", "TB", "LOZG", "FLM", "HFMS", "HFM", "KITE", "KMSC", "HEXA", "HRS", "BUT", "BULL", "GANE", "GUN", "CU", "RAD", "DECA", "UNQ", "ECA", "CMC", "OLMC", "OEC", "HSC", "BULT", "SSC", "MOVL", "PSC", "RMC", "ORC", "RRC", "OLMB"
]

ORDER_MAPPING = {
    'COLOR NAME': COLOR_ORDER,
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
    'TINC_NAME': INCLUSION_ORDER,
    'SINC_NAME': INCLUSION_ORDER,
}

DECIMAL_COLUMNS = ['Carat', 'RapRate', 'Reserve%', 'Rate', 'Amount', 'Ratio', 'Dep', 'Tab', 'Girdle', 'C/A', 'C/H', 'P/A', 'P/H', 'Disc%', 'Price/carat', 'Total Amount', 'Depth', 'Table']

DESIRED_COLUMN_SEQUENCE = [
    "Sr.No", "Tag_No", "Serial No", "DETAIL", "Shp", "Carat", "Color", "Qua", "Cut", "Pol", "Sym",
    "Flo", "Lab", "ReportNo", "RapRate", "Reserve%", "Disc%", "Rate", "Amount", "Measurement",
    "Dep", "Tab", "Ratio", "Shade", "Milky", "E.Facet", "TOIN_NAME", "SOIN_NAME", "GIRDLE_OPEN_NAME", "PAV_OPEN_NAME", "TBINC_NAME", "SBINC_NAME", "Remarks"
]

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
    "SGPComment": ["SGPComment"],
    "Dia": ["Dia"],
    "EYE Clean": ["EYE Clean"],
    "Luster": ["Luster"],
    "E.Facet": ["E.Facet"],
    "ShapeDesc": ["ShapeDesc"],
    "SUB_LOCATION": ["SUB_LOCATION"]
}

def get_carat_rank(carat_value):
    if not isinstance(carat_value, (int, float)): return 99
    if carat_value >= 15.00: return 1
    elif 8.00 <= carat_value <= 14.99: return 2
    elif 5.00 <= carat_value <= 7.99: return 3
    elif 4.00 <= carat_value <= 4.99: return 4
    elif 3.00 <= carat_value <= 3.99: return 5
    elif 2.50 <= carat_value <= 2.99: return 6
    elif 2.00 <= carat_value <= 2.49: return 7
    elif 1.70 <= carat_value <= 1.99: return 8
    elif 1.50 <= carat_value <= 1.69: return 9
    elif 1.20 <= carat_value <= 1.49: return 10
    elif 1.00 <= carat_value <= 1.19: return 11
    elif 0.90 <= carat_value <= 0.99: return 12
    elif 0.70 <= carat_value <= 0.89: return 13
    elif 0.50 <= carat_value <= 0.69: return 14
    elif 0.40 <= carat_value <= 0.49: return 15
    elif 0.01 <= carat_value <= 0.39: return 16
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

def apply_tolerance_filter(df, filter_rules, tolerance_level):
    if df.empty or filter_rules.empty: return pd.DataFrame()
    filtered_df = df.copy()
    conditions = []

    column_mapping = {
            'SHP': 'Shp', 'QUA': 'Qua', 'CUT': 'Cut', 'POL': 'Pol', 'SYM': 'Sym',
            'FLO': 'Flo', 'LAB': 'Lab', 'DISC%': 'Disc%', 'SHADE': 'Shade',    # 👈 Changed 'Reserve%' to 'Disc%' here
            'MILKY': 'Milky', 'E.FACET': 'E.Facet', 'TOIN_NAME': 'TOIN_NAME',
            'SOIN_NAME': 'SOIN_NAME', 'GIRDLE_OPEN_NAME': 'GIRDLE_OPEN_NAME',
            'PAV_OPEN_NAME': 'PAV_OPEN_NAME', 'TBINC_NAME': 'TBINC_NAME',
            'SBINC_NAME': 'SBINC_NAME', 'MEASUREMENT': 'Measurement',
            'REMARKS': 'Remarks', 'COLOR NAME': 'Color', 'CARAT': 'Carat',
            'RESERVE%': 'Disc%', 'RATIO': 'Ratio', 'RATE': 'Rate', 'RAPRATE': 'RapRate', # 👈 Changed 'Reserve%' to 'Disc%' here
            'AMOUNT': 'Amount', 'TABLE INCL': 'Table Incl', 'TINC_NAME': 'TINC_NAME', 'SINC_NAME': 'SINC_NAME', 'DEP': 'Dep', 'TAB': 'Tab'
    }
    remarks = str(filter_rules.get('REMARKS', 'INCLUDE')).strip().upper()

    for filter_col, filter_value in filter_rules.items():
        # Added TAG_NO here so it doesn't get treated as a filter condition
        if filter_col in ['SOLD PARTY', 'NORMALIZED PARTY', 'REMARKS', 'SALES_PERSON', 'TAG_NO'] or pd.isna(filter_value) or str(filter_value).strip() == '': continue
        stock_col = column_mapping.get(filter_col)
        if not stock_col or stock_col not in filtered_df.columns:
            if filter_col in INCLUSION_COLUMNS:
                raise ValueError(f'Missing stock column required by filter: {filter_col}')
            continue

        try:
            val_str = str(filter_value).strip().upper()
            if filter_col in INCLUSION_COLUMNS:
                conditions.append(inclusion_condition(filtered_df[stock_col], val_str, tolerance_level))
            elif filter_col in ['SHP', 'CARAT', 'LAB']:
                if filter_col == 'CARAT': cond = parse_carat_filter(filtered_df[stock_col], val_str)
                else:
                    acc = parse_filter_value(val_str, tolerance_level=0)
                    cond = filtered_df[stock_col].astype(str).str.upper().isin(acc) if acc else pd.Series(False, index=filtered_df.index)
                conditions.append(cond)
            elif filter_col in ['RESERVE%', 'DISC%', 'RATIO', 'RATE', 'RAPRATE', 'AMOUNT', 'DEP', 'TAB']:
                filtered_df[stock_col] = pd.to_numeric(filtered_df[stock_col], errors='coerce')
                conditions.append(parse_numeric_filter(filtered_df[stock_col], val_str))
            elif filter_col == 'MEASUREMENT':
                conditions.append(parse_measurement_filter(filtered_df[stock_col], val_str))
            elif filter_col in ORDER_MAPPING:
                order = ORDER_MAPPING[filter_col]
                apply_map = (ENABLE_FLO_MAPPING and filter_col == 'FLO')

                acc = parse_filter_value(val_str, order, tolerance_level, apply_flo_mapping=apply_map)

                series_to_check = filtered_df[stock_col].astype(str).str.upper()
                if apply_map:
                    series_to_check = series_to_check.replace(['VSL', 'SL'], 'FNT')

                if acc: conditions.append(series_to_check.isin(acc))
            else:
                acc = parse_filter_value(val_str, tolerance_level=0)
                if acc: conditions.append(filtered_df[stock_col].astype(str).str.upper().isin(acc))
        except:
            if filter_col in INCLUSION_COLUMNS:
                raise
            continue

    if not conditions: return pd.DataFrame()
    final_cond = pd.Series(True, index=filtered_df.index)
    for c in conditions: final_cond &= c

    if remarks == 'EXCLUDE': return filtered_df[~final_cond]
    else: return filtered_df[final_cond]

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

def auto_adjust_column_widths(worksheet):
    for column in worksheet.columns:
        col_letter = get_column_letter(column[0].column)
        col_name = column[0].value
        if col_name in ["CERTIFICATE", "VIDEO", "IMAGE", "DETAIL", "PHYGITAL"]:
            worksheet.column_dimensions[col_letter].width = 15
        else:
            max_length = max((len(str(cell.value)) for cell in column if cell.value is not None), default=0)
            worksheet.column_dimensions[col_letter].width = (max_length + 2) * 1.2

def apply_styles_to_worksheet(worksheet, dataframe=None):
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")

    exact_match_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
    nearby_match_fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
    maybe_match_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    headers = [cell.value for cell in worksheet[1]]

    if dataframe is not None and 'Match_Type' in dataframe.columns:
        for r_idx, row in enumerate(dataframe.itertuples(), 2):
            match_type = getattr(row, 'Match_Type', None)
            fill = None
            if match_type == 'Exact': fill = exact_match_fill
            elif match_type == 'Nearby': fill = nearby_match_fill
            elif match_type == 'Maybe': fill = maybe_match_fill
            if fill:
                for cell in worksheet[r_idx]: cell.fill = fill

    if worksheet.title == "Navigation":
        header_font = Font(bold=True, color="0000FF")
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
        for row in worksheet.iter_rows():
            for cell in row:
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        return

    for col_num, header in enumerate(headers, 1):
        if header in DECIMAL_COLUMNS:
            for cell in worksheet.iter_rows(min_row=2, min_col=col_num, max_col=col_num):
                if isinstance(cell[0].value, (int, float)):
                    cell[0].number_format = numbers.FORMAT_NUMBER_00

    if 'Sr.No' in headers:
        sr_no_col_idx = headers.index('Sr.No') + 1
        blue_underline_font = Font(color="0000FF", underline="single")
        for cell in worksheet.iter_rows(min_row=2, min_col=sr_no_col_idx, max_col=sr_no_col_idx):
            cell[0].font = blue_underline_font

    for row in worksheet.iter_rows():
        for cell in row:
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    if headers and headers[0] is not None and worksheet.title != "Navigation":
        for cell in worksheet[1]:
            cell.fill = header_fill

def process_party_data(stocks_df, party_filter_rule, tolerance_level):
    norm_party_name = party_filter_rule.get('NORMALIZED PARTY', '')
    party_display_name = party_filter_rule.get('SOLD PARTY', '')

    print(f"\n{'='*60}")
    print(f"PROCESSING PARTY: {party_display_name}")
    print(f"Tolerance: ±{tolerance_level}")

    # 1. Get the list of stones
    final_df = apply_tolerance_filter(stocks_df.copy(), party_filter_rule, tolerance_level)

    if not final_df.empty:
        # 2. Classify stones
        final_df['Match_Type'] = final_df.apply(lambda row: classify_match_type(row, party_filter_rule), axis=1)
        final_df['Is_Filter_Match'] = True

        print(f"    Matches: {len(final_df)} | Breakdown: "
              f"Exact={len(final_df[final_df['Match_Type']=='Exact'])} "
              f"Nearby={len(final_df[final_df['Match_Type']=='Nearby'])} "
              f"Maybe={len(final_df[final_df['Match_Type']=='Maybe'])}")

        final_df['Shp'] = pd.Categorical(final_df['Shp'], categories=SHAPE_ORDER, ordered=True)
        final_df['Color'] = pd.Categorical(final_df['Color'], categories=COLOR_ORDER, ordered=True)
        final_df['Qua'] = pd.Categorical(final_df['Qua'], categories=CLARITY_ORDER, ordered=True)
        final_df['Carat_Rank'] = final_df['Carat'].apply(get_carat_rank)

        match_type_order = ['Exact', 'Nearby', 'Maybe']
        final_df['Match_Type'] = pd.Categorical(final_df['Match_Type'], categories=match_type_order, ordered=True)

        final_df.sort_values(by=['Match_Type', 'Shp', 'Carat_Rank', 'Color', 'Qua'], ascending=[True, True, True, True, True], inplace=True)

        final_df.drop(columns=['Carat_Rank'], inplace=True)
        final_df['Shp'] = final_df['Shp'].astype(object)
        final_df['Color'] = final_df['Color'].astype(object)
        final_df['Qua'] = final_df['Qua'].astype(object)

        final_df.reset_index(drop=True, inplace=True)
        if 'Sr.No' not in final_df.columns: final_df.insert(0, 'Sr.No', range(1, 1 + len(final_df)))

        # --- TAG COLUMN INSERTION (FIXED) ---
        if ENABLE_RELATED_TAG:
            # We explicitly look for TAG_NO since that is what pandas converts 'Tag_No' to.
            # We also check 'TAG' just in case you ever change the column name.
            tag_val = party_filter_rule.get('TAG_NO', party_filter_rule.get('TAG', ''))
            final_df['Tag_No'] = '' if pd.isna(tag_val) else tag_val

        current_cols = final_df.columns.tolist()
        ordered_cols = []

        active_sequence = DESIRED_COLUMN_SEQUENCE.copy()
        if ENABLE_RELATED_TAG and 'Tag_No' not in active_sequence:
            active_sequence.append('Tag_No')

        for col in active_sequence:
            if col in current_cols: ordered_cols.append(col)
        for col in current_cols:
            if col not in ordered_cols: ordered_cols.append(col)
        final_df = final_df[ordered_cols]
    else:
        print(f"❌ NO MATCHES found for {party_display_name}")

    country = party_filter_rule.get('COUNTRY', '')
    city = party_filter_rule.get('CITY', '')
    return final_df, party_display_name, country, city

def load_filter_data(filter_file_path):
    removed_parties = []
    change_format_filters = pd.DataFrame()
    try:
        xl = pd.ExcelFile(filter_file_path)
        if "Removed Party" in xl.sheet_names:
            removed_parties_df = pd.read_excel(filter_file_path, sheet_name="Removed Party")
            if "Sold Party" in removed_parties_df.columns:
                removed_parties = removed_parties_df["Sold Party"].dropna().apply(normalize_party_name).tolist()
        if "Change Format" in xl.sheet_names:
            change_format_filters = pd.read_excel(filter_file_path, sheet_name="Change Format")
            change_format_filters.columns = change_format_filters.columns.str.strip().str.upper()
            if "SOLD PARTY" in change_format_filters.columns and "SALES_PERSON" in change_format_filters.columns:
                change_format_filters['NORMALIZED PARTY'] = change_format_filters['SOLD PARTY'].apply(normalize_party_name)
                change_format_filters['SALES_PERSON'] = change_format_filters['SALES_PERSON'].str.strip().str.title()
    except Exception as e:
        print(f"Error loading filter file {filter_file_path}: {str(e)}")
    return removed_parties, change_format_filters

def generate_salesperson_reports(file_path_stock, output_zip_path, filter_file_path, tolerance_level):
    print("🚀 STARTING REPORT GENERATION")
    stocks_df = pd.read_excel(file_path_stock, sheet_name='Sheet1')

    if 'Serial No' in stocks_df.columns: stocks_df.drop_duplicates(subset=['Serial No'], keep='first', inplace=True)
    reverse_stock_map = {}
    for final_name, variations in STOCK_COLUMN_MAPPING.items():
        for var in variations: reverse_stock_map[var.upper()] = final_name
    stocks_df.columns = [reverse_stock_map.get(col.strip().upper(), col.strip()) for col in stocks_df.columns]

    removed_parties, all_change_format_filters = load_filter_data(filter_file_path)
    salespersons_from_filters = all_change_format_filters['SALES_PERSON'].unique().tolist() if 'SALES_PERSON' in all_change_format_filters.columns else []

    if not salespersons_from_filters: return

    with zipfile.ZipFile(output_zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for salesperson in salespersons_from_filters:
            if pd.isna(salesperson): continue
            print(f"\n👤 PROCESSING SALESPERSON: {salesperson}")
            salesperson_party_reports = {}
            if 'NORMALIZED PARTY' in all_change_format_filters.columns:
                filter_parties = all_change_format_filters[(all_change_format_filters['SALES_PERSON'] == salesperson)]
                for index, party_rule in filter_parties.iterrows():
                    result, party_display_name, country, city = process_party_data(stocks_df, party_rule, tolerance_level)
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

                    if not all_results_for_party.empty:
                        if 'Sr.No' in all_results_for_party.columns: all_results_for_party.drop(columns=['Sr.No'], inplace=True)
                        all_results_for_party.insert(0, 'Sr.No', range(1, 1 + len(all_results_for_party)))

                        party_display_name = party_info['display_name']
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

                        # 1. Write Headers
                        for c_idx, col in enumerate(df_output.columns, 1):
                            ws.cell(row=1, column=c_idx, value=col)

                        # 2. Write Data Rows (without Excel string formulas)
                        for r_idx, row in df_output.iterrows():
                            for c_idx, col in enumerate(df_output.columns, 1):
                                ws.cell(row=r_idx + 2, column=c_idx, value=row[col])

                        # 3. Apply the native openpyxl hyperlink formatting
                        format_native_hyperlinks(ws, df_output)

                        # --------------------------------------------------

                        apply_styles_to_worksheet(ws, df_output)
                        auto_adjust_column_widths(ws)

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

                wb.save(buffer)
                clean_sp = "".join(x for x in str(salesperson) if x.isalnum() or x in (" ", "_")).strip()[:50] or "Report"
                fname = f"{clean_sp}.xlsx"
                zipf.writestr(fname, buffer.getvalue())
                print(f"✅ Saved report for {salesperson}")

    print(f"📦 Output file: {output_zip_path}")

if __name__ == "__main__":
    file_path_stock = "/content/drive/MyDrive/Filters File base Only Suggested Stones/stock Actual.xlsx"
    filter_file_path = "/content/drive/MyDrive/Filters File base Only Suggested Stones/filters.xlsx"
    output_zip_path = "/content/drive/MyDrive/Filters File base Only Suggested Stones/Regular_Stock_CC_Demand.zip"
    generate_salesperson_reports(file_path_stock, output_zip_path, filter_file_path, TOLERANCE_LEVEL)
