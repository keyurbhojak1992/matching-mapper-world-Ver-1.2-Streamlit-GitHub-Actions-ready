import pandas as pd
import zipfile
import os
from glob import glob
from io import BytesIO
import re
from tqdm.notebook import tqdm
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Border, Side, PatternFill, Font, numbers
from openpyxl import load_workbook, Workbook
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor, TwoCellAnchor
from openpyxl.worksheet.hyperlink import Hyperlink
from copy import copy
import tempfile
import shutil

# --- Global Constants ---
TOP_N_SUPERTAGS = 251

COLOR_ORDER = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O-P', 'Q-R', 'S-T', 'U-V']
CLARITY_ORDER = ['FL', 'IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2', 'SI3', 'I1', 'I2', 'I3', 'I4', 'I5']
FLUORESCENCE_ORDER = ['NON', 'VSL', 'SL', 'FNT', 'MED', 'STG', 'VST']
LAB_ORDER = ['IGI', 'GIA', 'HRD', 'NONE']

# List of columns to format to 2 decimal places
DECIMAL_COLUMNS = ['Carat', 'RapRate', 'Reserve%', 'Rate', 'Amount', 'Ratio', 'Dep', 'Tab', 'Dia', 'Girdle', 'C/A', 'C/H', 'P/A', 'P/H', 'Dis%', 'Price/carat', 'Total Amount', 'Depth', 'Table']

# Desired column sequence for the output files
DESIRED_COLUMN_SEQUENCE = [
    "Serial No", "DETAIL", "Shp", "Carat", "Color name", "Qua", "Cut", "Pol", "Sym",
    "Flo", "Lab", "ReportNo", "RapRate", "Reserve%", "Rate", "Amount", "Measurement",
    "Dep", "Tab", "Ratio"
]

# --- Custom Sorting Orders ---
SHAPE_ORDER = [
    "RBC", "PEAR", "PMB", "OVAL", "OMB", "MQ", "MMB", "EME", "SEM", "RMB", "RM", "ECB", "ECMB", "ECM", "ECBF", "CB", "CMB", "CM", "HRT", "HMB", "PR", "TR", "BGT", "PMC", "OMC", "MMC", "SRMB", "SRM", "TP", "SLD", "CAD", "SCAD", "TB", "LOZG", "FLM", "HFMS", "HFM", "KITE", "KMSC", "HEXA", "HRS", "BUT", "BULL", "GANE", "GUN", "CU", "RAD", "DECA", "UNQ", "ECA", "CMC", "OLMC", "OEC", "HSC", "BULT", "SSC", "MOVL", "PSC", "RMC", "BRIO", "FISH", "PF", "PRC", "MRC", "ORC", "RRC", "OLMB"
]
COLOR_FULL_ORDER = [
    "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O-P", "Q-R", "S-T", "U-V", "W-X", "Y-Z"
]
CLARITY_FULL_ORDER = [
    "FL", "IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2", "SI3", "I1", "I2", "I3", "I4", "I5"
]

# --- NEW ADDITIONS FOR PRANAV CHALODIYA ---

# Mapping for specific number formats, now including L, W, and H
COLUMN_FORMAT_MAPPING = {
    'Carat': numbers.FORMAT_NUMBER_00,
    'Reserve%': numbers.FORMAT_NUMBER_00,
    'Ratio': numbers.FORMAT_NUMBER_00,
    'C/A': numbers.FORMAT_NUMBER_00,
    'C/H': numbers.FORMAT_NUMBER_00,
    'P/A': numbers.FORMAT_NUMBER_00,
    'P/H': numbers.FORMAT_NUMBER_00,
    'Girdle': numbers.FORMAT_NUMBER_00,
    'Dep': '0.0',
    'RapRate': numbers.FORMAT_NUMBER,
    'Rate': numbers.FORMAT_NUMBER,
    'Amount': numbers.FORMAT_NUMBER,
    'Tab': numbers.FORMAT_NUMBER,
    'Dis%': numbers.FORMAT_NUMBER_00,
    'Price/carat': numbers.FORMAT_NUMBER_00,
    'Total Amount': numbers.FORMAT_NUMBER_00,
    'Depth': numbers.FORMAT_NUMBER_00,
    'Table': numbers.FORMAT_NUMBER_00,
    'L': numbers.FORMAT_NUMBER_00,
    'W': numbers.FORMAT_NUMBER_00,
    'H': numbers.FORMAT_NUMBER_00,
}

# Column name mapping for Mr. Pranav Chalodiya
PRANAV_COLUMN_MAPPING = {
    "Loc": ["Location"],
    "Sr.NO": ["Sr.NO"],
    "Serial No": ["Serial No"],
    "DETAIL": ["DETAIL"],
    "Shp": ["Shp", "Shape", "SHAPE"],
    "Carat": ["Carat", "CARAT"],
    "Color": ["Color name", "Color"],
    "Qua": ["Qua", "Clarity", "Cla"],
    "Cut": ["Cut", "cut"],
    "Pol": ["Pol", "polish"],
    "Sym": ["Sym"],
    "Flo": ["Flo", "Flou."],
    "Lab": ["Lab"],
    "RapRate": ["RapRate", "Rap"],
    "Reserve%": ["Reserve%", "Disc%", "Dis%"],
    "Rate": ["Rate", "Price/Ct$"],
    "Amount": ["Amount", "Amount$"],
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
    "Table Incl": ["Table Incl", "Table Inclusion"],
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
    "ReportNo": ["ReportNo"],
    "SGPComment": ["SGPComment"],
    "Dia": ["Dia"],
    "EYE Clean": ["EYE Clean"],
    "Luster": ["Luster"],
    "E.Facet": ["E.Facet"],
    "ShapeDesc": ["ShapeDesc"],
    "SUB_LOCATION": ["SUB_LOCATION"]
}

SALESPERSON_COLUMN_SEQUENCES = {
    "Mr. Pranav Chalodiya": [
        "Loc", "Sr.NO", "Serial No", "DETAIL", "Shp", "Carat", "Color", "Qua", "Cut", "Pol", "Sym",
        "Flo", "Lab", "RapRate", "Reserve%", "Rate", "Amount", "Milky", "Shade", "L", "W", "H",
        "Dep", "Tab", "Ratio", "KeyToSymbol", "comment", "Table Incl", "TOIN_NAME", "TBINC_NAME",
        "SINC_NAME", "SOIN_NAME", "SBINC_NAME", "GIRDLE_OPEN_NAME", "C/A", "C/H",
        "P/A", "P/H"
    ]
}

STRICT_COLUMN_SALESPERSONS = {
    "Mr. Pranav Chalodiya"
}

MAPPING_TABLES = {
    "Table Incl": {
        "NONE": "NON", "MIN PINPOINT": "PP1", "MED PINPOINT": "PP2", "MIN CRYSTAL": "CRL1", "MED CRYSTAL": "CRL2", "MIN SPREADED": "SPR1", "MED SPREADED": "SPR2", "HVY SPREADED": "SPR3", "MIN FEATHER": "FR1", "MED FEATHER": "FR2", "HVY FEATHER": "FR3", "MIN CLOUD": "CLD1", "MED CLOUD": "CLD2", "HVY CLOUD": "CLD3", "HVY CRYSTAL": "CRL3", "MINOR PINPOINT": "PP1", "MEDIUM PINPOINT": "PP2", "MINOR CRYSTAL": "CRL1", "MEDIUM CRYSTAL": "CRL2", "MINOR SPREADED": "SPR1", "MEDIUM SPREADED": "SPR2", "MINOR FEATHER": "FR1", "MEDIUM FEATHER": "FR2", "HEAVY FEATHER": "FR3", "MINOR CLOUD": "CLD1", "MEDIUM CLOUD": "CLD1"
    },
    "SINC_NAME": {
        "NONE": "NON", "MIN PINPOINT": "PP1", "MED PINPOINT": "PP2", "MIN CRYSTAL": "CRL1", "MED CRYSTAL": "CRL2", "MIN SPREADED": "SPR1", "MED SPREADED": "SPR2", "HVY SPREADED": "SPR3", "MIN FEATHER": "FR1", "MED FEATHER": "FR2", "HVY FEATHER": "FR3", "MIN CLOUD": "CLD1", "MED CLOUD": "CLD2", "HVY CLOUD": "CLD3", "HVY CRYSTAL": "CRL3", "MINOR PINPOINT": "PP1", "MEDIUM PINPOINT": "PP2", "MINOR CRYSTAL": "CRL1", "MEDIUM CRYSTAL": "CRL2", "MINOR SPREADED": "SPR1", "MEDIUM SPREADED": "SPR2", "MINOR FEATHER": "FR1", "MEDIUM FEATHER": "FR2", "MINOR CLOUD": "CLD1", "MEDIUM CLOUD": "CLD1"
    },
    "Shade": {
        "NONE": "NON", "LIGHT BROWN": "LB", "BROWN": "LB1", "STRONG BROWN": "LB2", "VERY STRONG BROWN": "LB3", "LIGHT GREEN": "GR1", "GREEN": "GR2", "STRONG GREEN": "GR3", "VERY STRONG GREEN": "GR4", "MIX TINGE 1": "MT", "MIX TINGE 2": "MT1", "MIX TINGE 3": "MT2", "INSPECTION ADVISABLE": "IA"
    },
    "Milky": {
        "NONE": "NON", "LIGHT MILKY": "LM", "MILKY": "LM1", "STRONG MILKY": "LM2", "VERY STRONG MILKY": "LM3", "INSPECTION ADVISABLE": "IA"
    },
    "TBINC_NAME": {
        "NONE": "NON", "PIN POINT BLK": "TBP1", "MINOR BLK": "TB1", "NOTICEABLE BLACK": "TBN", "MEDIUM BLK": "TB2", "HEAVY BLK": "TB3"
    },
    "SBINC_NAME": {
        "NONE": "NON", "PIN POINT BLK": "SBP1", "MINOR BLK": "SB1", "NOTICEABLE BLACK": "SBN", "MEDIUM BLK": "SB2", "HEAVY BLK": "SB3"
    },
    "TOIN_NAME": {
        "NONE": "NON", "HAIR LINE": "THL", "SMALL OPEN": "TSO", "MEDIUM OPEN": "TMO", "HVY OPEN": "THO"
    },
    "SOIN_NAME": {
        "NONE": "NON", "HAIR LINE": "SHL", "SMALL OPEN": "SSO", "MEDIUM OPEN": "SMO", "HVY OPEN": "SHO"
    },
    "PAV_OPEN_NAME": {
        "NONE": "NON", "HAIR LINE": "PHL", "SMALL OPEN": "PSO", "MEDIUM OPEN": "PMO", "HVY OPEN": "PHO"
    },
    "GIRDLE_OPEN_NAME": {
        "NONE": "NON", "HAIR LINE": "GHL", "SMALL OPEN": "GSO", "MEDIUM OPEN": "GMO", "HVY OPEN": "GHO"
    }
}


# --- Helper Functions ---

def normalize_party_name(name):
    """Normalizes a party name by removing special characters, converting to uppercase, and stripping common business suffixes."""
    if not isinstance(name, str):
        return ""
    name = re.sub(r'[^\w\s]', '', name).upper().strip()
    name = re.sub(r'\b(INC|LLC|LTD|CO|CORP|COMPANY)\b', '', name).strip()
    return name

def get_tolerance_values(value, order):
    """Gets values within a tolerance range (previous, current, next) from an ordered list."""
    try:
        index = order.index(value)
        start = max(0, index - 1)
        end = min(len(order) - 1, index + 1)
        return order[start:end + 1]
    except ValueError:
        return [value]

def load_sales_data(sales_files):
    """
    Loads sales data from multiple Excel files, renames columns, creates a 'Super Tag',
    and normalizes 'Sold Party' and 'Salesperson' names."""
    chunks = []
    for file in sales_files:
        try:
            sheet_names_to_try = ['Sale', 'Sheet1', 'Sheet 1', 'Data', 'Sales']
            df = None
            xl = pd.ExcelFile(file)
            available_sheets = xl.sheet_names
            for sheet_name in sheet_names_to_try:
                if sheet_name in available_sheets:
                    df = pd.read_excel(file, sheet_name=sheet_name)
                    break
            if df is None and available_sheets:
                df = pd.read_excel(file, sheet_name=available_sheets[0])
            if df is not None:
                df.columns = df.columns.str.strip().str.upper()
                required_cols = {'SHP', 'CARAT_GRP', 'COL', 'QUA', 'FLO', 'LAB', 'SOLD PARTY', 'SALE_PERSON'}
                available_cols = {col for col in df.columns}
                missing_cols = required_cols - available_cols

                if not missing_cols:
                    df = df.rename(columns={
                        'SHP': 'Shp', 'CARAT_GRP': 'Carat_grp', 'COL': 'Col', 'QUA': 'Qua',
                        'FLO': 'Flo', 'LAB': 'Lab', 'SOLD PARTY': 'Sold Party', 'SALE_PERSON': 'Salesperson'
                    })
                    df = df[['Shp', 'Carat_grp', 'Col', 'Qua', 'Flo', 'Lab', 'Sold Party', 'Salesperson']].copy()
                    df['Super Tag'] = df[['Shp', 'Carat_grp', 'Col', 'Qua', 'Flo', 'Lab']].astype(str).agg('-'.join, axis=1)
                    chunks.append(df)
                else:
                    print(f"Skipping file {file} due to missing required columns: {missing_cols}")
        except Exception as e:
            print(f"Error processing file {file}: {str(e)}")

    if chunks:
        combined = pd.concat(chunks, ignore_index=True)
        combined['Normalized Party'] = combined['Sold Party'].apply(normalize_party_name)
        combined['Salesperson'] = combined['Salesperson'].str.strip().str.title()
        return combined
    return pd.DataFrame()

def auto_adjust_column_widths(worksheet):
    """Adjusts column widths in an openpyxl worksheet based on content."""
    for column in worksheet.columns:
        col_letter = get_column_letter(column[0].column)
        col_name = worksheet.cell(row=1, column=column[0].column).value

        if col_name in ["CERTIFICATE", "VIDEO", "IMAGE", "DETAIL", "PHYGITAL"]:
            worksheet.column_dimensions[col_letter].width = 15
        else:
            max_length = 0
            for cell in column:
                if cell.value is not None:
                    try:
                        cell_len = len(str(cell.value))
                        if cell_len > max_length:
                            max_length = cell_len
                    except TypeError:
                        pass
            worksheet.column_dimensions[col_letter].width = (max_length + 2) * 1.2

def apply_styles_to_worksheet(worksheet, salesperson):
    """
    Applies general styles (borders, alignment) to cells and specific formats to columns,
    including custom formatting for specific salespersons."""
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    exact_match_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
    hyperlink_font = Font(color="0000FF", underline="single")

    headers = [cell.value for cell in worksheet[1]]

    # Freeze the top row and first two columns
    worksheet.freeze_panes = 'C2'

    # Corrected addition: Find the Reserve% column and make it bold
    try:
        reserve_col_idx = headers.index('Reserve%') + 1
        for row_idx in range(1, worksheet.max_row + 1):
            cell = worksheet.cell(row=row_idx, column=reserve_col_idx)
            current_font = cell.font
            cell.font = Font(name=current_font.name, size=current_font.size, bold=True, italic=current_font.italic, vertAlign=current_font.vertAlign, underline=current_font.underline, strike=current_font.strike, color=current_font.color)
    except ValueError:
        pass # The column does not exist, so do nothing

    if 'Is_Nearby' in headers:
        exact_match_col_idx = headers.index('Is_Nearby') + 1
        for row_idx in range(2, worksheet.max_row + 1):
            cell = worksheet.cell(row=row_idx, column=exact_match_col_idx)
            if cell.value is not None and str(cell.value).lower() == 'false':
                for data_cell in worksheet[row_idx]:
                    data_cell.fill = exact_match_fill

    if salesperson == "Mr. Pranav Chalodiya":
        for col_num, header in enumerate(headers, 1):
            if header in COLUMN_FORMAT_MAPPING:
                for row_idx in range(2, worksheet.max_row + 1):
                    cell = worksheet.cell(row=row_idx, column=col_num)
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = COLUMN_FORMAT_MAPPING[header]
    else:
        for col_num, header in enumerate(headers, 1):
            if header in DECIMAL_COLUMNS:
                for row_idx in range(2, worksheet.max_row + 1):
                    cell = worksheet.cell(row=row_idx, column=col_num)
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = numbers.FORMAT_NUMBER_00

    for row in worksheet.iter_rows():
                for cell in row:
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    # FIX: Check for native hyperlink property instead of formula string
                    if cell.hyperlink:
                        cell.font = hyperlink_font

    if headers and headers[0] is not None and worksheet.title != "Navigation":
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = Font(bold=True)

def process_party_data(party_group, stocks_df, top_n):
    """
    Processes sales data for a specific party, finding exact and tolerance matches
    in the stock data based on 'Super Tag'."""
    party_super_tags = party_group['Super Tag'].value_counts().head(top_n)
    top_n_tags = party_super_tags.index.tolist()

    exact_matches_list = []
    tolerance_suggestions_list = []
    seen_serials = set()

    for tag in top_n_tags:
        tag_parts = tag.split('-')
        if len(tag_parts) < 6:
            continue

        col = tag_parts[-4]
        qua = tag_parts[-3]
        flo = tag_parts[-2]
        lab = tag_parts[-1]
        shp = tag_parts[0]
        carat_grp = '-'.join(tag_parts[1:-4])

        exact_matches = stocks_df[
            (stocks_df['Shp'] == shp) &
            (stocks_df['Carat_grp'] == carat_grp) &
            (stocks_df['Color'] == col) &
            (stocks_df['Qua'] == qua) &
            (stocks_df['Flo'] == flo) &
            (stocks_df['Lab'] == lab)
        ].copy()

        if not exact_matches.empty:
            exact_matches = exact_matches[~exact_matches['Serial No'].isin(seen_serials)]
            if not exact_matches.empty:
                seen_serials.update(exact_matches['Serial No'].tolist())
                exact_matches['Is_Nearby'] = False
                exact_matches_list.append(exact_matches)

        color_range = get_tolerance_values(col, COLOR_ORDER)
        clarity_range = get_tolerance_values(qua, CLARITY_ORDER)
        fluorescence_range = get_tolerance_values(flo, FLUORESCENCE_ORDER)

        tolerance_matches = stocks_df[
            (stocks_df['Shp'] == shp) &
            (stocks_df['Carat_grp'] == carat_grp) &
            (stocks_df['Color'].isin(color_range)) &
            (stocks_df['Qua'].isin(clarity_range)) &
            (stocks_df['Flo'].isin(fluorescence_range)) &
            (stocks_df['Lab'] == lab) &
            (~stocks_df['Serial No'].isin(seen_serials))
        ].copy()

        if not tolerance_matches.empty:
            seen_serials.update(tolerance_matches['Serial No'].tolist())
            tolerance_matches['Is_Nearby'] = True
            tolerance_suggestions_list.append(tolerance_matches)

    combined_result = pd.DataFrame()
    if exact_matches_list:
        combined_result = pd.concat(exact_matches_list)
    if tolerance_suggestions_list:
        combined_result = pd.concat([combined_result] + tolerance_suggestions_list) if not combined_result.empty else pd.concat(tolerance_suggestions_list)

    if not combined_result.empty:
        combined_result['Lab_order'] = pd.Categorical(combined_result['Lab'], categories=LAB_ORDER, ordered=True)
        combined_result['Shp_order'] = pd.Categorical(combined_result['Shp'], categories=SHAPE_ORDER, ordered=True)
        combined_result['Color_order'] = pd.Categorical(combined_result['Color'], categories=COLOR_FULL_ORDER, ordered=True)
        combined_result['Qua_order'] = pd.Categorical(combined_result['Qua'], categories=CLARITY_FULL_ORDER, ordered=True)
        combined_result['Flo_order'] = pd.Categorical(combined_result['Flo'], categories=FLUORESCENCE_ORDER, ordered=True)

        combined_result = combined_result.sort_values(
            by=['Lab_order', 'Shp_order', 'Carat', 'Color_order', 'Qua_order', 'Flo_order'],
            ascending=[True, True, False, True, True, True]
        )

        combined_result.drop(
            columns=['Lab_order', 'Shp_order', 'Color_order', 'Qua_order', 'Flo_order'],
            inplace=True
        )
        combined_result.drop(columns=[col for col in ['Super Tag', 'Carat_grp'] if col in combined_result.columns], inplace=True)
        combined_result.reset_index(drop=True, inplace=True)

        current_cols = combined_result.columns.tolist()
        ordered_cols = []
        for col in DESIRED_COLUMN_SEQUENCE:
            if col in current_cols:
                ordered_cols.append(col)
        for col in current_cols:
            if col not in ordered_cols:
                ordered_cols.append(col)
        combined_result = combined_result[ordered_cols]

    return combined_result

def generate_salesperson_reports(file_path_stock, sales_files_pattern, output_zip_path, top_n):
    """
    Generates an Excel report for each salesperson, containing recommended stones
    based on their past sales. Each report is saved within a single ZIP file."""
    sales_data_dir = os.path.dirname(sales_files_pattern)
    if sales_data_dir and not os.path.exists(sales_data_dir):
        os.makedirs(sales_data_dir)
        print(f"Created directory for sales files: {sales_data_dir}")

    sales_files = glob(sales_files_pattern)
    if not sales_files:
        print(f"No sales files found matching pattern: {sales_files_pattern}. Please ensure sales files are uploaded.")
        return

    print(f"Processing {len(sales_files)} sales files...")
    print("Loading stock data...")
    try:
        stocks_df = pd.read_excel(file_path_stock, sheet_name='Sheet1')
        required_stock_cols = ['Shp', 'Carat_grp', 'Color', 'Qua', 'Flo', 'Lab']

        if "Mr. Pranav Chalodiya" in SALESPERSON_COLUMN_SEQUENCES:
            pranav_cols = set(SALESPERSON_COLUMN_SEQUENCES["Mr. Pranav Chalodiya"])
            for original_col, aliases in PRANAV_COLUMN_MAPPING.items():
                for alias in aliases:
                    if alias in stocks_df.columns:
                        stocks_df.rename(columns={alias: original_col}, inplace=True)
                        break

        if not all(col in stocks_df.columns for col in required_stock_cols):
            print(f"Error: Missing one or more required columns in stock file: {required_stock_cols}")
            return

        stocks_df['Super Tag'] = stocks_df[required_stock_cols].astype(str).agg('-'.join, axis=1)
        print(f"Loaded {len(stocks_df)} stock items")
    except FileNotFoundError:
        print(f"Error: Stock file not found at '{file_path_stock}'. Please check the path and upload the file.")
        return
    except Exception as e:
        print(f"Error loading stock data: {e}")
        return

    print("Loading sales data...")
    all_sales = load_sales_data(sales_files)
    if all_sales.empty:
        print("No valid sales data found in the files after processing.")
        return
    print(f"Loaded {len(all_sales)} sales records")

    salesperson_groups = all_sales.groupby('Salesperson')
    print(f"Found {len(salesperson_groups)} unique salespersons...")

    output_zip_dir = os.path.dirname(output_zip_path)
    if output_zip_dir and not os.path.exists(output_zip_dir):
        os.makedirs(output_zip_dir)
        print(f"Created output directory for zip file: {output_zip_dir}")

    with zipfile.ZipFile(output_zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for salesperson, salesperson_data in tqdm(salesperson_groups, desc="Processing salespersons"):
            party_groups = salesperson_data.groupby('Normalized Party')
            clean_salesperson = "".join(x for x in salesperson if x.isalnum() or x in (" ", "_", "-")).strip()
            salesperson_dir = f"{clean_salesperson}/"

            for norm_party, party_group in party_groups:
                result = process_party_data(party_group, stocks_df, top_n)

                if not result.empty:
                    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temp_excel_file:
                        temp_file_name = temp_excel_file.name

                        wb = Workbook()
                        ws = wb.active
                        ws.title = "Suggested_Stones"

                        most_common_party_display_name = party_group['Sold Party'].mode()[0]

                        if salesperson == "Mr. Pranav Chalodiya":
                            reverse_pranav_mapping = {alias: preferred_name for preferred_name, aliases in PRANAV_COLUMN_MAPPING.items() for alias in aliases}
                            df_output = result.rename(columns=reverse_pranav_mapping).copy()
                            for col, mapping in MAPPING_TABLES.items():
                                if col in df_output.columns:
                                    df_output[col] = df_output[col].replace(mapping)
                        else:
                            df_output = result.copy()

                        if salesperson in SALESPERSON_COLUMN_SEQUENCES:
                            ordered_cols = SALESPERSON_COLUMN_SEQUENCES[salesperson]
                            df_output = df_output.loc[:, df_output.columns.intersection(ordered_cols)]
                            df_output = df_output.reindex(columns=ordered_cols)
                            if 'Sr.NO' in df_output.columns:
                                df_output['Sr.NO'] = range(1, len(df_output) + 1)
                        else:
                            current_cols = df_output.columns.tolist()
                            ordered_cols = []
                            for col in DESIRED_COLUMN_SEQUENCE:
                                if col in current_cols:
                                    ordered_cols.append(col)
                            for col in current_cols:
                                if col not in ordered_cols:
                                    ordered_cols.append(col)
                            df_output = df_output[ordered_cols]

                        # 1. Write the Headers
                        headers = df_output.columns.tolist()
                        for c_idx, col_name in enumerate(headers, 1):
                            ws.cell(row=1, column=c_idx, value=col_name)

                        # 2. Write the Data and Native Hyperlinks
                        hyperlink_cols = {"CERTIFICATE", "VIDEO", "IMAGE", "DETAIL", "PHYGITAL"}

                        for r_idx, row in enumerate(df_output.itertuples(index=False), 2):
                            for c_idx, value in enumerate(row, 1):
                                col_name = headers[c_idx - 1]
                                cell = ws.cell(row=r_idx, column=c_idx)

                                # FIX: Apply native openpyxl hyperlink to bypass 255 character limit
                                if col_name in hyperlink_cols and pd.notna(value) and isinstance(value, str) and 'http' in value.lower():
                                    cell.value = col_name      # Set the display text (e.g., "CERTIFICATE")
                                    cell.hyperlink = value     # Attach the long URL as metadata
                                else:
                                    cell.value = value         # Standard data entry

                        auto_adjust_column_widths(ws)
                        apply_styles_to_worksheet(ws, salesperson)

                        if "Is_Nearby" in df_output.columns:
                            current_ws_headers = [cell.value for cell in ws[1]]
                            try:
                                col_idx_to_delete_from_ws = current_ws_headers.index("Is_Nearby") + 1
                                ws.delete_cols(col_idx_to_delete_from_ws)
                            except ValueError:
                                pass

                        wb.save(temp_file_name)

                        # --- FIX APPLIED HERE: Added "&" to the allowed characters ---
                        clean_party_name = re.sub(r'[\\/*?:"<>|]', "", most_common_party_display_name).strip()

                        if len(clean_party_name) > 50:
                            clean_party_name = clean_party_name[:50].strip()
                        if not clean_party_name:
                            clean_party_name = f"PartyReport_{norm_party[:10]}"

                        filename = f"{clean_party_name}.xlsx"

                        zipf.write(temp_file_name, arcname=f"{salesperson_dir}{filename}")
                        print(f"Saved report for '{most_common_party_display_name}' as '{salesperson_dir}{filename}' inside the zip.")
                        os.remove(temp_file_name)

    print(f"\nAll reports saved to '{output_zip_path}'")

# --- Main Execution Block ---
if __name__ == "__main__":
    print("Running script.")
    print("- 'stock Actual.xlsx' should be in the Colab root directory (`/content/`).")
    print("- Your sales Excel files should be in a folder named 'sales_data' inside the Colab root directory (`/content/sales_data/`).")
    print("  For example: `/content/sales_data/sale_report_jan.xlsx`, `/content/sales_data/sale_report_feb.xlsx`")

    #stock_file = "/content/stock Actual Pranav n.xlsx"
    #sales_files_pattern = "/content/sales_data/*.xlsx"
    #output_zip = "/content/suggested_Bid_stones_Last_Bid_Regular_sales_Base_Pranav_17.03.2026.zip"

    stock_file = "/content/drive/MyDrive/Pranav Suggested Files/Stock Actual Pranav n.xlsx"
    sales_files_pattern = "/content/drive/MyDrive/Pranav Suggested Files/sales_data/*.xlsx"
    output_zip = "/content/drive/MyDrive/Pranav Suggested Files/suggested_Bid_stones_Last_Bid_Regular_sales_Base_USA_12-08-2026_Check.zip"

    print(f"\nStock file path: {stock_file}")
    print(f"Sales files pattern: {sales_files_pattern}")
    print(f"Output ZIP path: {output_zip}")

    generate_salesperson_reports(
        file_path_stock=stock_file,
        sales_files_pattern=sales_files_pattern,
        output_zip_path=output_zip,
        top_n=TOP_N_SUPERTAGS
    )