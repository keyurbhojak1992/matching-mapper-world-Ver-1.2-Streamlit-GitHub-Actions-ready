import pandas as pd
import zipfile
import os
import glob
from io import BytesIO
import re
import math
from collections import defaultdict
from copy import copy
from datetime import datetime
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Border, Side, PatternFill, Font, numbers
from openpyxl import Workbook
import numpy as np

# =============================================================================
# --- USER INPUT SECTION ---
# =============================================================================
TOP_N_SUPERTAGS = None
TOLERANCE_LEVEL = 1
ENABLE_STYLING = True

# --- FLUORESCENCE CONTROL ---
ENABLE_FLO_MAPPING = True

# --- MATCHING CONFIGURATION (MUST MATCH PROFILE CREATION) ---
INCLUDE_COLOR         = True
INCLUDE_CLARITY       = True
INCLUDE_FLORO         = True

INCLUDE_CUT           = False
INCLUDE_POLISH        = False
INCLUDE_SYMMETRY      = False

INCLUDE_SHADE         = False
INCLUDE_MILKY         = False

INCLUDE_TABLE_BLACK   = False
INCLUDE_SIDE_BLACK    = False

INCLUDE_TABLE_OPEN    = False
INCLUDE_SIDE_OPEN     = False
INCLUDE_GIRDLE_OPEN   = False
INCLUDE_PAVILION_OPEN = False

# --- TOLERANCE CONTROLS ---
ENABLE_COLOR_TOLERANCE         = True
ENABLE_CLARITY_TOLERANCE       = True
ENABLE_FLORO_TOLERANCE         = True

ENABLE_CUT_TOLERANCE           = False
ENABLE_POLISH_TOLERANCE        = True
ENABLE_SYMMETRY_TOLERANCE      = True

ENABLE_SHADE_TOLERANCE         = True
ENABLE_MILKY_TOLERANCE         = True

ENABLE_TABLE_BLACK_TOLERANCE   = True
ENABLE_SIDE_BLACK_TOLERANCE    = True

ENABLE_TABLE_OPEN_TOLERANCE    = True
ENABLE_SIDE_OPEN_TOLERANCE     = True
ENABLE_GIRDLE_OPEN_TOLERANCE   = True
ENABLE_PAVILION_OPEN_TOLERANCE = True

# --- SALES SUPERTAG DISPLAY CONTROLS ---
SHOW_BASIC_SUPERTAG    = True
SHOW_ADVANCED_SUPERTAG = False

# --- REQUESTED FEATURES ---
SHOW_REF_STONE_ID = True
SHOW_CONCAT_STATUS = False

# --- OUTPUT FORMAT TOGGLES ---
GENERATE_INDIVIDUAL_REPORTS = False
GENERATE_MERGED_MASTER = True
ENABLE_CLEAN_FILE = True
ENABLE_PAIR_MODE = True

ENABLE_TYPE_COLUMNS = False
STATIC_TYPE_SEQUENCE = ["Type_SALE", "Type_BID", "Type_SEARCH", "Type_WISHLIST", "Type_DETAILS", "Type_VIEW", "Type_CART", "Type_OFFER"]
ENABLE_SALES_SUPERTAG_COLUMN = False

MAX_ROWS_PER_SHEET = 1000000

TEST_MODE = True
# BID Suggested
#TEST_SALESPEOPLE = ["Mr. Bhavik Hadiya", "Mr. Dhaval Moradiya", "Mr. Parva Shah", "Mr. Hiren Rathod", "Mr. Jagdish Tarsariya", "Mr. Shubham Patil", "Mr. Aakash Parmar", "Mr. Aatish Sarodiya", "Mr. Amit Tarsaria", "Mr. Kalash Sanghavi", "Mr. Mukul Shah", "Mr. Nikunj Shankar", "Mr. Shubham Tarsariya", "Mr. Sunil Bhatiya"]

#Regular Suggested
#TEST_SALESPEOPLE = ["Mr. Bhavik Hadiya", "Mr. Dhaval Moradiya", "Mr. Parva Shah", "Mr. Hiren Rathod", "Mr. Jagdish Tarsariya", "Mr. Shubham Patil", "Mr. Aakash Parmar", "Mr. Aatish Sarodiya", "Mr. Amit Tarsaria", "Mr. Hrithik Hundia", "Mr. Kalash Sanghavi", "Mr. Mukul Shah", "Mr. Nikunj Shankar", "Mr. Shubham Tarsariya", "Mr. Sunil Bhatiya", "Ms. Zeal Harsora", "Mr. Ridham Lakhani", "Mr. Pranav Chalodiya"]
#TEST_SALESPEOPLE = ["Mr. Bhavik Hadiya", "Mr. Dhaval Moradiya", "Mr. Parva Shah", "Mr. Hiren Rathod", "Mr. Jagdish Tarsariya", "Mr. Shubham Patil", "Mr. Nikunj Shankar", "Mr. Shubham Tarsariya"]

## USE For Master Merge file
TEST_SALESPEOPLE = ["Mr. Bhavik Hadiya", "Mr. Dhaval Moradiya", "Mr. Parva Shah", "Mr. Hiren Rathod", "Mr. Jagdish Tarsariya", "Mr. Shubham Patil", "Mr. Aakash Parmar", "Mr. Aatish Sarodiya", "Mr. Amit Tarsaria", "Mr. Hrithik Hundia", "Mr. Kalash Sanghavi", "Mr. Mukul Shah", "Mr. Nikunj Shankar", "Mr. Shubham Tarsariya", "Mr. Sunil Bhatiya", "Ms. Zeal Harsora", "Mr. Ridham Lakhani", "Mr. Pranav Chalodiya"]

#Single Specific User
#TEST_SALESPEOPLE = ["Ms. Zeal Harsora"]

TEST_MODE_PARTY = False
#TEST_PARTIES_filter = ["AGARWAL DIAMONDS MANUFACTURING PVT LTD", "AGARWAL DIAMONDS", "ANTWERP SALES INTERNATIONAL INC", "AVNER EIZENSTIGAN", "B&L DIAMONDS LTD", "B.NAMDAR DIAMONDS LTD", "BULCHI INTERNATIONAL", "DIAM STAR", "ERAN DIAMONDS-YEHUDA SAYAG DIAMONDS LTD", "Gopal Gems", "H.R.DIAMONDS PVT LTD", "ISHTARA DIAMONDS B.V", "MOSHE NAMDAR", "NAKODIAM DIAMONDS PVT. LTD.", "NOGA DIAMONDS LTD", "OFER MIZRAHI DIAMONDS INC", "PAR DIAM", "PARISHI DIAMOND INC", "PINALL EXPORT", "RK GEMS", "YARO TRADING", "ZOYA & CO."]
#TEST_PARTIES = ["B.NAMDAR DIAMONDS LTD", "BRIGHTON DIAMOND CO., LTD.", "GID DIAMONDS", "GREEN HILLS DIAMOND BROKERS", "H&Y CREATION BANGKOK CO LTD", "L & R JOSYFON LTD", "NOGA DIAMONDS LTD", "RENADIAM",  "SPECTRUM OVERSEAS DMCC.", "SUPER SOLITAIRE HK LTD", "UTSAV DIAMOND"]
#TEST_PARTIES = ["NS LUXURY", "VIDHYA GEMS", "NOOR JEWELS", "AL BUTA JEWELLERY LLC", "MELEK JEWELS EXPORTS", "FACETS INDIA", "AARAV IMPEX", "DIAMARS FINE JEWELS LLP", "R.H.EXPO", "SHRADDHA GEMS", "CARRERA JEWELLERY", "PRIORITY JEWELS LIMITED"]
TEST_PARTIES = ["ABHAY GEMS"]

# --- OUTPUT FILTER CONTROL NEW ---
MATCH_STATUS_FILTER_MODE = "NONE" # "ALL"NONE
MATCH_STATUS_ALLOWED_TYPES = ["Exact", "Near by", "Filter"]
MATCH_STATUS_SPECIFIC_PARTIES = ["NONE"]

ENABLE_NEW_FILL_LOGIC = True
ENABLE_TOP_CUSTOMERS_HIGHLIGHT = True
TOP_CUSTOMERS_N = 10

# =============================================================================
# --- OUTPUT SPLIT & CLEANUP CONTROLS (NEW) ---
# =============================================================================
ENABLE_SMALL_OUTPUT_DELETE = False
SMALL_OUTPUT_THRESHOLD = 5

## Split Excel Output File ##
ENABLE_OUTPUT_FILE_SEPARATED = False

# =============================================================================
# --- ADVANCED SORTING & LIMITING CONTROL ---
# =============================================================================
ENABLE_ADVANCED_SORTING = False

SORT_COUNT_THRESHOLD = 250
SORT_NORMAL_PARTY_LIMIT = 120
SORT_HIGH_PRIORITY_PARTY_LIMIT = 150

SORT_EXACT_CAP = 3
SORT_NEAR_CAP = 2
SORT_MAYBE_CAP = 1

SORT_HIGH_PRIORITY_PARTY_TYPES = {"PLATINUM", "GOLD", "DIAMOND", "PREMIUM", "VIP"}
SORT_HIGH_PRIORITY_TOTAL_HISTORY_THRESHOLD = 200

# =============================================================================
# --- MASTER TAG STRUCTURE ---
# =============================================================================
FULL_TAG_MAP = [
    'Shp', 'Carat_grp', 'Color', 'Qua', 'Flo', 'Lab',
    'Cut', 'Pol', 'Sym',
    'Shade', 'Milky',
    'TBINC_NAME', 'SBINC_NAME',
    'TOIN_NAME', 'SOIN_NAME', 'GIRDLE_OPEN_NAME', 'PAV_OPEN_NAME'
]

BASIC_TAG_MAP = ['Shp', 'Carat_grp', 'Color', 'Qua', 'Flo', 'Lab']

# --- CONSTANTS ---
COLOR_ORDER = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O-P', 'Q-R', 'S-T', 'U-V', 'W-X', 'Y-Z']
CLARITY_ORDER = ['FL', 'IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2', 'SI3', 'I1', 'I2', 'I3', 'I4', 'I5', 'I6', 'I7', 'I8']

if ENABLE_FLO_MAPPING:
    FLUORESCENCE_ORDER = ['NON', 'FNT', 'MED', 'STG', 'VST']
else:
    FLUORESCENCE_ORDER = ['NON', 'VSL', 'SL', 'FNT', 'MED', 'STG', 'VST']

CUT_POL_SYM_ORDER = ["F", "G", "VG", "EX"]

# --- ZERO-SPACE ARRAYS FIX ---
SHADE_ORDER = ['NONE', 'LIGHTBROWN', 'BROWN', 'STRONGBROWN', 'VERYSTRONGBROWN', 'LIGHTGREEN', 'GREEN', 'STRONGGREEN', 'VERYSTRONGGREEN', 'MIXTINGE1', 'MIXTINGE2', 'INSPECTIONADVISABLE', 'MIXTINGE3']
MILKY_ORDER = ['NONE', 'LIGHTMILKY', 'MILKY', 'STRONGMILKY', 'VERYSTRONGMILKY', 'INSPECTIONADVISABLE', 'IACLOUD', 'LIGHTCLOUD', 'CLOUD']
E_FACET_ORDER = ['NON', 'PS', 'CS', 'PM', 'CM', 'TEF']
TOIN_SOIN_GIRDLE_PAV_OPEN_ORDER = ['NONE', 'HAIRLINE', 'SMALLOPEN', 'MEDIUMOPEN', 'HEAVYOPEN', 'HVYOPEN']
TBINC_SBINC_ORDER = ['NONE', 'PINPOINTBLK', 'MINORBLK', 'NOTICEABLEBLACK', 'MEDIUMBLK', 'HEAVYBLK']
TB_INC_SB_INC_ORDER = ['NONE', 'MINORPINPOINT', 'MEDIUMPINPOINT', 'MINORCRYSTAL', 'MEDIUMCRYSTAL', 'MINORSPREADED', 'MEDIUMSPREADED', 'HEAVYSPREADED', 'MINORFEATHER', 'MEDIUMFEATHER', 'HEAVYFEATHER', 'MINORCLOUD', 'MEDIUMCLOUD', 'HEAVYCLOUD', 'HEAVYCRYSTAL']

SHAPE_ORDER = [
    "RBC", "PEAR", "PMB", "OVAL", "OMB", "MQ", "MMB", "EME", "SEM", "RMB", "RM", "ECB",
    "ECMB", "ECM", "ECBF", "CB", "CMB", "CM", "HRT", "HMB", "PR", "TR", "BGT", "PMC",
    "OMC", "MMC", "SRMB", "SRM", "TP", "SLD", "CAD", "SCAD", "TB", "LOZG", "FLM",
    "HFMS", "HFM", "KITE", "KMSC", "HEXA", "HRS", "BUT", "BULL", "GANE", "GUN", "CU",
    "RAD", "DECA", "UNQ", "ECA", "CMC", "OLMC", "OEC", "HSC", "BULT", "SSC", "MOVL",
    "PSC", "RMC", "ORC", "RRC", "OLMB"]

STD_CARAT_ORDER = [
    "0.30-0.39", "0.40-0.49", "0.50-0.69", "0.70-0.89", "0.90-0.99",
    "1.00-1.19", "1.20-1.49", "1.50-1.69", "1.70-1.99", "2.00-2.49",
    "2.50-2.99", "3.00-3.99", "4.00-4.99", "5.00-7.99", "8.00-14.99", "15+"]

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

HK_COUNTRIES = ['HONGKONG', 'HK', 'H.K', 'CHINA']

ORDER_MAPPING = {
    'COLOR NAME': COLOR_ORDER, 'Color': COLOR_ORDER,
    'QUA': CLARITY_ORDER, 'Qua': CLARITY_ORDER,
    'FLO': FLUORESCENCE_ORDER, 'Flo': FLUORESCENCE_ORDER,
    'CUT': CUT_POL_SYM_ORDER, 'Cut': CUT_POL_SYM_ORDER,
    'POL': CUT_POL_SYM_ORDER, 'Pol': CUT_POL_SYM_ORDER,
    'SYM': CUT_POL_SYM_ORDER, 'Sym': CUT_POL_SYM_ORDER,
    'SHADE': SHADE_ORDER, 'Shade': SHADE_ORDER,
    'MILKY': MILKY_ORDER, 'Milky': MILKY_ORDER,
    'E.FACET': E_FACET_ORDER,
    'TOIN_NAME': TOIN_SOIN_GIRDLE_PAV_OPEN_ORDER, 'SOIN_NAME': TOIN_SOIN_GIRDLE_PAV_OPEN_ORDER,
    'GIRDLE_OPEN_NAME': TOIN_SOIN_GIRDLE_PAV_OPEN_ORDER, 'PAV_OPEN_NAME': TOIN_SOIN_GIRDLE_PAV_OPEN_ORDER,
    'TBINC_NAME': TBINC_SBINC_ORDER, 'SBINC_NAME': TBINC_SBINC_ORDER,
    'TABLE INCL': TB_INC_SB_INC_ORDER, 'SINC_NAME': TB_INC_SB_INC_ORDER
}

DECIMAL_COLUMNS = ['Carat', 'RapRate', 'Reserve%', 'Rate', 'Amount', 'Ratio', 'Dep', 'Tab',
                   'Dia', 'Girdle', 'C/A', 'C/H', 'P/A', 'P/H', 'Disc%', 'Price/carat',
                   'Total Amount', 'Depth', 'Table']

DESIRED_COLUMN_SEQUENCE = [
    "Sr.No", "Serial No", "DETAIL", "PHYGITAL", "Shp", "Carat", "Color", "Qua", "Cut", "Pol", "Sym",
    "Flo", "Lab", "ReportNo", "RapRate", "Reserve%", "Disc%", "Rate", "Amount", "SGPComment", "Loc", "Measurement",
    "Ratio", "Dep", "Tab", "Dia", "Girdle", "C/A", "C/H", "P/A", "P/H", "EYE Clean", "Luster", "E.Facet", "Culet", "TINC_NAME",
    "SINC_NAME", "TOIN_NAME", "SOIN_NAME", "GIRDLE_OPEN_NAME", "PAV_OPEN_NAME", "TBINC_NAME", "SBINC_NAME",
    "Spot Incl.", "KeyToSymbol", "Comment", "Remarks", "IMAGE", "VIDEO", "CERTIFICATE", "Shade", "Milky"]

DESIRED_COLUMN_SEQUENCE_AAKASH = [
    "Sr.No", "Serial No", "DETAIL", "PHYGITAL", "Shp", "Carat", "Color", "Qua", "Cut", "Pol", "Sym",
    "Flo", "Lab", "ReportNo", "RapRate", "Reserve%", "Disc%", "Rate", "Amount", "SGPComment", "Loc", "Measurement",
    "Dep", "Tab", "Ratio", "Dia", "Girdle", "C/A", "C/H", "P/A", "P/H", "EYE Clean", "Luster", "E.Facet", "Culet", "TINC_NAME",
    "SINC_NAME", "TOIN_NAME", "SOIN_NAME", "GIRDLE_OPEN_NAME", "PAV_OPEN_NAME", "TBINC_NAME", "SBINC_NAME", "Spot Incl.", "KeyToSymbol", "Comment", "Remarks"]

STOCK_COLUMN_MAPPING = {
    "Loc": ["Location"], "Sr.NO": ["Sr.NO"], "Serial No": ["Serial No", "Stone ID"],
    "DETAIL": ["DETAIL"], "Shp": ["Shp", "Shape"], "Carat": ["Carat"],
    "Color": ["Color name", "Color"], "Qua": ["Qua", "Clarity"], "Cut": ["Cut", "cut"],
    "Pol": ["Pol", "polish"], "Sym": ["Sym"], "Flo": ["Flo", "Flou."], "Lab": ["Lab"],
    "RapRate": ["RapRate", "Rap"], "Disc%": ["Reserve%", "Disc%", "Dis%"],
    "Rate": ["Rate", "Price/Ct$", "$/Cts.", "Price/carat"], "Amount": ["Total Amount", "Amount", "Amount$", "Total"], "Milky": ["Milky"],
    "Shade": ["Shade"], "L": ["L", "Length"], "W": ["W", "Width"], "H": ["H", "Height"],
    "Dep": ["Dep", "Depth%", "Depth"], "Tab": ["Tab", "Table%", "Table"], "Ratio": ["Ratio", "Ratio"],
    "KeyToSymbol": ["KeyToSymbol", "Key To Symbol"],
    "comment": ["comment", "Comment"], "Measurement": ["Measurement", "Measurment"],
    "TINC_NAME": ["Table Incl", "Table Inclusion"], "TOIN_NAME": ["TOIN_NAME", "Table Open"],
    "TBINC_NAME": ["TBINC_NAME", "Table Black", "Table Black Incl."], "SINC_NAME": ["SINC_NAME", "Side Inclusion"],
    "SOIN_NAME": ["SOIN_NAME", "Side Open"], "SBINC_NAME": ["SBINC_NAME", "Side Black", "Side Black Incl."],
    "GIRDLE_OPEN_NAME": ["GIRDLE_OPEN_NAME", "Girdle Open"], "PAV_OPEN_NAME": ["Pavilion Open"], "C/A": ["C/A", "CR Angle"],
    "C/H": ["C/H", "CR Height"], "P/A": ["P/A", "Pav Angle"], "P/H": ["P/H", "Pav Height"],
    "Culet": ["Culet"], "Girdle": ["Girdle", "Girdle %"], "IMAGE": ["IMAGE"],
    "VIDEO": ["VIDEO"], "CERTIFICATE": ["CERTIFICATE"], "PHYGITAL": ["PHYGITAL"],
    "ReportNo": ["ReportNo", "Certificate No"],
    "SGPComment": ["SGPComment", "SGP COMMENT"], "Dia": ["Dia"],
    "EYE Clean": ["EYE Clean"], "Luster": ["Luster"], "E.Facet": ["E.Facet"],
    "ShapeDesc": ["ShapeDesc"], "LOC": ["SUB_LOCATION"], "Comment": ["comment", "Comment"],
    "Color": ["Color name", "Color"], "Loc": ["Loc."]}

# =============================================================================
# --- ADVANCED SORTING MODULE ---
# =============================================================================

def sort_norm(v):
    if pd.isna(v) or v is None: return ""
    return str(v).strip()

def sort_norm_upper(v):
    return sort_norm(v).upper()

def sort_safe_float(v, default=0.0):
    try:
        if pd.isna(v) or v is None: return default
        if isinstance(v, float) and math.isnan(v): return default
        return float(v)
    except Exception:
        return default

def get_pair_id_for_sorting(text):
    if pd.isna(text): return None
    match = re.search(r'(?:PAIRED\s+WITH|MATCH\s+PAIR)\s*[^\d]*([\d\.]+)', str(text), re.IGNORECASE)
    if match:
        s = str(match.group(1)).strip()
        if s.endswith('.0'): s = s[:-2]
        elif s.endswith('.00'): s = s[:-3]
        return s
    return None

def cps_score(cut, pol, sym):
    rank = {"EX": 4, "VG": 3, "G": 2, "F": 1}
    return rank.get(sort_norm_upper(cut), 0) + rank.get(sort_norm_upper(pol), 0) + rank.get(sort_norm_upper(sym), 0)

def build_family_key(shape, carat_grp, color, qua, lab):
    return "|".join([sort_norm_upper(shape), sort_norm(carat_grp), sort_norm_upper(color), sort_norm_upper(qua), sort_norm_upper(lab)])

def build_variant_key(cut, pol, sym, flo):
    return "|".join([sort_norm_upper(cut), sort_norm_upper(pol), sort_norm_upper(sym), sort_norm_upper(flo)])

def derive_carat_group(carat_value, parsed_ranges):
    try: carat = float(carat_value)
    except Exception: return ""
    for low, high, label in parsed_ranges:
        if low <= carat <= high: return label
    if parsed_ranges:
        nearest = min(parsed_ranges, key=lambda x: min(abs(carat - x[0]), abs(carat - x[1])))
        return nearest[2]
    return ""

def load_sales_history_for_sorting(sales_path):
    all_sales_dfs = []
    if os.path.isdir(sales_path):
        print(f"  -> Scanning folder for sales files: {sales_path}")
        file_paths = [f for f in glob.glob(os.path.join(sales_path, "*.xlsx")) if not os.path.basename(f).startswith('~$')]
        for fp in file_paths:
            try:
                df = pd.read_excel(fp)
                all_sales_dfs.append(df)
            except Exception as e:
                pass
    elif os.path.isfile(sales_path) and sales_path.endswith('.xlsx'):
        try:
            all_sales_dfs.append(pd.read_excel(sales_path))
        except Exception as e:
            pass

    if not all_sales_dfs:
        print("  -> WARNING: No sales history files found. Sorting will bypass history limits.")
        return None

    sales = pd.concat(all_sales_dfs, ignore_index=True)
    sales.columns = [str(c).strip() for c in sales.columns]

    ranges = []
    if "Carat_grp" in sales.columns:
        seen = set()
        for grp in sales["Carat_grp"].dropna().astype(str).unique():
            grp = grp.strip()
            if grp in seen: continue
            seen.add(grp)
            m = re.match(r"^\s*(\d+(?:\.\d+)?)\s*-\s*(\d+(?:\.\d+)?)\s*$", grp)
            if m: ranges.append((float(m.group(1)), float(m.group(2)), grp))
        ranges.sort(key=lambda x: (x[0], x[1]))

    if "Type" not in sales.columns: sales["Type"] = ""

    type_weights = {"SALE": 20, "BID": 15, "OFFER": 10, "WISHLIST": 10, "CART": 10, "HOLD": 7, "INQUIRY": 2, "ENQUIRY": 2, "LIST": 2, "VIEW": 4, "SEARCH": 3, "DETAILS": 5}
    sales["TYPE_WEIGHT"] = sales["Type"].apply(lambda x: type_weights.get(sort_norm_upper(x), 1))

    sales["FAMILY_KEY"] = sales.apply(lambda r: build_family_key(r.get("Shp"), r.get("Carat_grp"), r.get("Col"), r.get("Qua"), r.get("Lab")), axis=1)

    party_family_map = defaultdict(float)
    global_family_map = defaultdict(float)
    party_total_map = defaultdict(float)
    party_type_map = {}

    for _, r in sales.iterrows():
        sold_party = sort_norm_upper(r.get("Sold Party"))
        family_key = r.get("FAMILY_KEY", "")
        wt = sort_safe_float(r.get("TYPE_WEIGHT"), 1)

        party_family_map[(sold_party, family_key)] += wt
        global_family_map[family_key] += wt
        party_total_map[sold_party] += wt

        if "Party_type" in sales.columns and sold_party and sold_party not in party_type_map:
            party_type_map[sold_party] = sort_norm_upper(r.get("Party_type"))

    return {"parsed_ranges": ranges, "party_family_map": dict(party_family_map), "global_family_map": dict(global_family_map), "party_total_map": dict(party_total_map), "party_type_map": dict(party_type_map)}

def get_priority_for_party(party_name, sales_info):
    party_u = sort_norm_upper(party_name)
    party_type = sales_info["party_type_map"].get(party_u, "")
    total_history = sales_info["party_total_map"].get(party_u, 0)
    if party_type in SORT_HIGH_PRIORITY_PARTY_TYPES: return "HIGH"
    if total_history >= SORT_HIGH_PRIORITY_TOTAL_HISTORY_THRESHOLD: return "HIGH"
    return "NORMAL"

def is_clean_row(shade, milky, sgp_comment):
    shade = sort_norm_upper(shade)
    milky = sort_norm_upper(milky)
    sgp = sort_norm_upper(sgp_comment)
    bad_words = ["I.A", "OPEN", "SHADE", "MILKY", "CLOUD", "BLACK", "INCLUSION"]
    shade_ok = shade in ("", "NONE", "NO", "NIL")
    milky_ok = milky in ("", "NONE", "NO", "NIL")
    sgp_bad = any(w in sgp for w in bad_words)
    return 1 if (shade_ok and milky_ok and not sgp_bad) else 0

def score_tuple(item):
    match_rank = {"Exact": 3, "Near by": 2, "May be": 1}
    return (
        item["party_family_count"],               # PRIORITY 1: Sales History Weight
        match_rank.get(item["match_type"], 0),    # PRIORITY 2: Exact vs Near by
        item["is_confirmed_pair"],                # PRIORITY 3: Is it a verified pair with an ID?
        item["global_family_count"],
        item["is_clean"],
        item["cps"],
        item["disc"]
    )

def unique_keep_allowed(item):
    if item["match_type"] == "Exact": return True
    if item["is_confirmed_pair"]: return True
    if item["party_family_count"] >= 1: return True
    if item["global_family_count"] >= 5: return True
    return False

def select_keep_rows_stable(items, party_limit):
    families = defaultdict(list)
    for item in items: families[item["family_key"]].append(item)

    keep_row_nums = set()
    for fam_items in families.values():
        fam_sorted = sorted(fam_items, key=score_tuple, reverse=True)
        if len(fam_sorted) == 1:
            if unique_keep_allowed(fam_sorted[0]): keep_row_nums.add(fam_sorted[0]["row_num"])
            continue

        by_type = {"Exact": [], "Near by": [], "May be": []}
        for it in fam_sorted: by_type[it["match_type"]].append(it)

        for cap, t_list in [(SORT_EXACT_CAP, by_type["Exact"]), (SORT_NEAR_CAP, by_type["Near by"]), (SORT_MAYBE_CAP, by_type["May be"])]:
            chosen, seen_vars = 0, set()
            for it in t_list:
                if chosen >= cap: break
                if it["is_confirmed_pair"] or it["variant_key"] not in seen_vars:
                    keep_row_nums.add(it["row_num"])
                    seen_vars.add(it["variant_key"])
                    chosen += 1

    kept_items = [it for it in items if it["row_num"] in keep_row_nums]
    if len(kept_items) > party_limit:
        allowed = {it["row_num"] for it in sorted(kept_items, key=score_tuple, reverse=True)[:party_limit]}
        kept_items = [it for it in items if it["row_num"] in allowed]

    kept_items = sorted(kept_items, key=lambda x: x["row_num"])
    return [it["row_num"] for it in kept_items]

def apply_advanced_sorting_on_dataframe(df, party_name, sales_info):
    if df.empty: return df

    items = []
    records = df.to_dict('records')
    for i, row in enumerate(records):
        shp = str(row.get("Shp", ""))
        carat = str(row.get("Carat", ""))
        color = str(row.get("Color", ""))
        qua = str(row.get("Qua", ""))
        cut = str(row.get("Cut", ""))
        pol = str(row.get("Pol", ""))
        sym = str(row.get("Sym", ""))
        flo = str(row.get("Flo", ""))
        lab = str(row.get("Lab", ""))
        disc = str(row.get("Disc%", ""))
        sgp = str(row.get("SGPComment", ""))
        shade = str(row.get("Shade", ""))
        milky = str(row.get("Milky", ""))

        is_filter = row.get("Is_Filter_Match", False)
        tol_type = str(row.get("Tolerance_Type", "")).upper()
        is_nearby = row.get("Is_Nearby", False)

        if is_filter:
            match_type = "Exact"
        elif tol_type == "NEARBY" or (is_nearby and not ENABLE_NEW_FILL_LOGIC):
            match_type = "Near by"
        elif tol_type == "MAYBE":
            match_type = "May be"
        else:
            match_type = "Exact"

        carat_grp = derive_carat_group(carat, sales_info["parsed_ranges"])
        family_key = build_family_key(shp, carat_grp, color, qua, lab)
        variant_key = build_variant_key(cut, pol, sym, flo)

        party_family_count = sales_info["party_family_map"].get((sort_norm_upper(party_name), family_key), 0)
        global_family_count = sales_info["global_family_map"].get(family_key, 0)

        extracted_pair_id = get_pair_id_for_sorting(sgp)

        item = {
            "row_num": i,
            "match_type": match_type,
            "is_confirmed_pair": 1 if extracted_pair_id else 0,
            "party_family_count": sort_safe_float(party_family_count),
            "global_family_count": sort_safe_float(global_family_count),
            "cps": cps_score(cut, pol, sym),
            "disc": sort_safe_float(disc),
            "family_key": family_key,
            "variant_key": variant_key,
            "is_clean": is_clean_row(shade, milky, sgp)
        }
        items.append(item)

    party_limit = SORT_HIGH_PRIORITY_PARTY_LIMIT if get_priority_for_party(party_name, sales_info) == "HIGH" else SORT_NORMAL_PARTY_LIMIT

    keep_indices = select_keep_rows_stable(items, party_limit)
    return df.iloc[keep_indices].copy()

# =============================================================================
# --- MAIN APPLICATION LOGIC ---
# =============================================================================

def get_desired_sequence_for_salesperson(salesperson):
    if not isinstance(salesperson, str): return DESIRED_COLUMN_SEQUENCE
    key = salesperson.strip().lower()
    if "aakash" in key and "parmar" in key: return DESIRED_COLUMN_SEQUENCE_AAKASH
    return DESIRED_COLUMN_SEQUENCE

def global_clean_serial(val):
    if pd.isna(val) or val == "": return ""
    s = str(val).strip()
    if s.endswith('.0'): s = s[:-2]
    elif s.endswith('.00'): s = s[:-3]
    return s

def normalize_party_name(name):
    if not isinstance(name, str): return ""
    name = re.sub(r'[^\w\s]', '', name).upper().strip()
    name = re.sub(r'\s+', ' ', name).strip()
    return name

def get_carat_range_string(carat_value, country=None):
    if pd.isna(carat_value) or not isinstance(carat_value, (int, float)): return "UNKNOWN"
    val = float(carat_value)
    use_hk_logic = False
    if country and isinstance(country, str):
        if country.upper().strip() in HK_COUNTRIES: use_hk_logic = True
    if use_hk_logic:
        for (low, high, label) in HK_CHINA_RANGES:
            if low <= val <= high: return label
        return "5.00-99.99" if val >= 5.00 else "UNKNOWN"
    else:
        if 0.00 <= val <= 0.39: return "0.30-0.39"
        elif 0.40 <= val <= 0.49: return "0.40-0.49"
        elif 0.50 <= val <= 0.69: return "0.50-0.69"
        elif 0.70 <= val <= 0.89: return "0.70-0.89"
        elif 0.90 <= val <= 0.99: return "0.90-0.99"
        elif 1.00 <= val <= 1.19: return "1.00-1.19"
        elif 1.20 <= val <= 1.49: return "1.20-1.49"
        elif 1.50 <= val <= 1.69: return "1.50-1.69"
        elif 1.70 <= val <= 1.99: return "1.70-1.99"
        elif 2.00 <= val <= 2.49: return "2.00-2.49"
        elif 2.50 <= val <= 2.99: return "2.50-2.99"
        elif 3.00 <= val <= 3.99: return "3.00-3.99"
        elif 4.00 <= val <= 4.99: return "4.00-4.99"
        elif 5.00 <= val <= 7.99: return "5.00-7.99"
        elif 8.00 <= val <= 14.99: return "8.00-14.99"
        elif val >= 15.00: return "15+"
        return "UNKNOWN"

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
    elif 0.30 <= carat_value <= 0.39: return 16
    return 99

def get_tolerance_values(value, order, tolerance_level=1):
    try:
        index = order.index(value)
        start = max(0, index - tolerance_level)
        end = min(len(order) - 1, index + tolerance_level)
        return order[start:end + 1]
    except ValueError:
        return [value]

def get_neighbor_carat_grps(carat_grp):
    try:
        if carat_grp not in STD_CARAT_ORDER: return [carat_grp]
        idx = STD_CARAT_ORDER.index(carat_grp)
        start = max(0, idx - 1)
        end = min(len(STD_CARAT_ORDER) - 1, idx + 1)
        return STD_CARAT_ORDER[start:end + 1]
    except ValueError:
        return [carat_grp]

def parse_filter_value(filter_string, order_list=None, apply_flo_mapping=False):
    if pd.isna(filter_string) or not isinstance(filter_string, str): return []
    filter_string = filter_string.strip().upper()

    if apply_flo_mapping and order_list == FLUORESCENCE_ORDER:
        filter_string = filter_string.replace('VSL', 'FNT').replace('SL', 'FNT')

    # AUTO-CORRECT COMMON TYPOS
    filter_string = filter_string.replace("NOTICABLE", "NOTICEABLE")

    values = []
    if order_list and 'TO' in filter_string:
        parts = filter_string.split('TO')
        if len(parts) == 2:
            # STRIP SPACES FROM BOUNDARIES
            start_val = parts[0].replace(" ", "").strip()
            end_val = parts[1].replace(" ", "").strip()

            clean_order = [str(x).replace(" ", "").upper() for x in order_list]
            try:
                start_idx = clean_order.index(start_val)
                end_idx = clean_order.index(end_val)
                if start_idx <= end_idx: values.extend(order_list[start_idx:end_idx + 1])
                else: values.extend(order_list[end_idx:start_idx + 1])
            except ValueError: values.append(filter_string)
        else: values.append(filter_string)
    elif ',' in filter_string:
        for val in filter_string.split(','):
            values.append(val.strip())
    else:
        values.append(filter_string)
    return list(set(values))

def load_filter_data(filter_file_path):
    removed_parties = []
    change_format_filters = pd.DataFrame()
    tolerance_df = pd.DataFrame()

    print(f"\n--- Loading Filter File: {os.path.basename(filter_file_path)} ---")

    if not os.path.exists(filter_file_path):
        print(f"❌ ERROR: Filter file DOES NOT EXIST at {filter_file_path}")
        return removed_parties, change_format_filters, tolerance_df

    try:
        xl = pd.ExcelFile(filter_file_path)
        if "Removed Party" in xl.sheet_names:
            removed_parties_df = pd.read_excel(filter_file_path, sheet_name="Removed Party")
            if "Sold Party" in removed_parties_df.columns:
                removed_parties = removed_parties_df["Sold Party"].dropna().apply(normalize_party_name).tolist()
            print(f"  ✅ 'Removed Party' sheet: Loaded {len(removed_parties)} rules.")

        if "Change Format" in xl.sheet_names:
            change_format_filters = pd.read_excel(filter_file_path, sheet_name="Change Format")
            change_format_filters.columns = change_format_filters.columns.str.strip().str.upper()
            if "SOLD PARTY" in change_format_filters.columns:
                change_format_filters['NORMALIZED PARTY'] = change_format_filters['SOLD PARTY'].apply(normalize_party_name)
            if "SALES_PERSON" in change_format_filters.columns:
                change_format_filters['SALES_PERSON'] = change_format_filters['SALES_PERSON'].astype(str).str.strip().str.title()
            print(f"  ✅ 'Change Format' sheet: Loaded {len(change_format_filters)} advanced rules.")

        if "Tolerance" in xl.sheet_names:
            tolerance_df = pd.read_excel(filter_file_path, sheet_name="Tolerance")
            tolerance_df.columns = tolerance_df.columns.str.strip().str.upper()
            if "SOLD PARTY" in tolerance_df.columns:
                tolerance_df['NORMALIZED PARTY'] = tolerance_df['SOLD PARTY'].apply(normalize_party_name)
            if "SALES_PERSON" in tolerance_df.columns:
                tolerance_df['SALES_PERSON'] = tolerance_df['SALES_PERSON'].astype(str).str.strip().str.title()
            print(f"  ✅ 'Tolerance' sheet: Loaded {len(tolerance_df)} tolerance rules.")

    except Exception as e:
        print(f"❌ CRITICAL ERROR reading {filter_file_path}: {e}")

    return removed_parties, change_format_filters, tolerance_df

def load_exclusion_data(clean_file_path):
    exclusion_df = pd.DataFrame()

    print(f"--- Loading Clean (No Need) File: {os.path.basename(clean_file_path)} ---")

    try:
        if not os.path.exists(clean_file_path):
            print(f"❌ WARNING: Clean file DOES NOT EXIST at {clean_file_path}")
            return exclusion_df

        exclusion_df = pd.read_excel(clean_file_path)
        exclusion_df.columns = exclusion_df.columns.str.strip().str.upper()

        if "SOLD PARTY" in exclusion_df.columns:
            exclusion_df['NORMALIZED PARTY'] = exclusion_df['SOLD PARTY'].apply(normalize_party_name)

        if "SALES_PERSON" in exclusion_df.columns:
            exclusion_df['SALES_PERSON'] = exclusion_df['SALES_PERSON'].astype(str).str.strip().str.title()

        print(f"  ✅ Clean file loaded: {len(exclusion_df)} exclusion rules found.\n")

    except Exception as e:
        print(f"❌ CRITICAL ERROR reading {clean_file_path}: {e}\n")

    return exclusion_df

def parse_numeric_filter(series, filter_string):
    filter_string = str(filter_string).strip().upper()
    try:
        if 'TO' in filter_string:
            parts = filter_string.split('TO')
            return (series >= float(parts[0])) & (series <= float(parts[1]))
        elif filter_string.startswith('>='): return series >= float(filter_string[2:])
        elif filter_string.startswith('>'): return series > float(filter_string[1:])
        elif filter_string.startswith('<='): return series <= float(filter_string[2:])
        elif filter_string.startswith('<'): return series < float(filter_string[1:])
        elif filter_string.startswith('!='): return series != float(filter_string[2:])
        else: return series == float(filter_string)
    except:
        return pd.Series(False, index=series.index)

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

def apply_advanced_filters(df, filter_rules):
    if df.empty or filter_rules.empty: return pd.DataFrame()
    filtered_df = df
    conditions = []

    column_mapping = {
        'SHP': 'Shp', 'QUA': 'Qua', 'CUT': 'Cut', 'POL': 'Pol', 'SYM': 'Sym',
        'FLO': 'Flo', 'LAB': 'Lab', 'DISC%': 'Disc%', 'SHADE': 'Shade', 'MILKY': 'Milky',
        'E.FACET': 'E.Facet', 'TOIN_NAME': 'TOIN_NAME', 'SOIN_NAME': 'SOIN_NAME',
        'GIRDLE_OPEN_NAME': 'GIRDLE_OPEN_NAME', 'PAV_OPEN_NAME': 'PAV_OPEN_NAME',
        'TBINC_NAME': 'TBINC_NAME', 'SBINC_NAME': 'SBINC_NAME', 'MEASUREMENT': 'Measurement',
        'REMARKS': 'Remarks', 'COLOR NAME': 'Color', 'CARAT': 'Carat', 'RESERVE%': 'Reserve%',
        'RATIO': 'Ratio', 'RATE': 'Rate', 'RAPRATE': 'RapRate', 'AMOUNT': 'Amount',
        'TABLE INCL': 'Table Incl', 'TAB': 'Tab', 'C/A': 'C/A', 'DEP': 'Dep', 'C/H': 'C/H',
        'P/A': 'P/A', 'P/H': 'P/H', 'GIRDLE': 'Girdle', 'DIA': 'Dia',
        'SGP COMMENT': 'SGPComment', 'SGPCOMMENT': 'SGPComment',
        'KEYTOSYMBOL': 'KeyToSymbol', 'COMMENT': 'Comment', 'COLOR DESC': 'Color Desc',
        'EYE CLEAN': 'EYE Clean', 'LUSTER': 'Luster', 'CULET': 'Culet'
    }

    LONG_TEXT_COLS = ['SGP COMMENT', 'SGPCOMMENT', 'KEYTOSYMBOL', 'COMMENT', 'COLOR DESC']

    remarks = str(filter_rules.get('REMARKS', 'INCLUDE')).strip().upper()

    for filter_col, filter_value in filter_rules.items():
        if filter_col in ['SOLD PARTY', 'NORMALIZED PARTY', 'REMARKS', 'SALES_PERSON',
                          'PAST_SALES', 'CARAT_GRP_TOLERANCE', 'TOLERANCE +-1', 'TOLERANCE +-2',
                          'COUNTRY', 'CITY'] \
           or pd.isna(filter_value) or str(filter_value).strip() == '':
            continue

        stock_col = column_mapping.get(filter_col)

        if not stock_col or stock_col not in filtered_df.columns:
            matched_col = next((c for c in filtered_df.columns if c.lower() == filter_col.lower()), None)
            if matched_col:
                stock_col = matched_col
            else:
                continue

        try:
            val_str = str(filter_value).strip().upper()
            if filter_col in ['CARAT', 'RESERVE%', 'DISC%', 'RATIO', 'RATE', 'RAPRATE',
                              'AMOUNT', 'TAB', 'C/A', 'DEP', 'C/H', 'P/A', 'P/H', 'GIRDLE', 'DIA']:
                cond = parse_carat_filter(filtered_df[stock_col], val_str) if filter_col == 'CARAT' else parse_numeric_filter(filtered_df[stock_col], val_str)
                conditions.append(cond)
            elif filter_col == 'MEASUREMENT':
                pass

            elif filter_col in LONG_TEXT_COLS:
                parts = [p.strip() for p in val_str.split(',')]
                inc = [p for p in parts if not p.startswith('!')]
                exc = [p.lstrip('!') for p in parts if p.startswith('!')]

                series_to_check = filtered_df[stock_col].fillna('').astype(str)
                col_conds = []

                if inc:
                    pattern_inc = '|'.join([rf"(?<!\w){re.escape(v)}(?!\w)" for v in inc if v])
                    if pattern_inc:
                        match_mask = series_to_check.str.contains(pattern_inc, case=False, regex=True, na=False)
                        col_conds.append(match_mask)

                if exc:
                    pattern_exc = '|'.join([rf"(?<!\w){re.escape(v)}(?!\w)" for v in exc if v])
                    if pattern_exc:
                        match_mask_exc = series_to_check.str.contains(pattern_exc, case=False, regex=True, na=False)
                        col_conds.append(~match_mask_exc)

                if col_conds:
                    comb = pd.Series(True, index=filtered_df.index)
                    for c in col_conds: comb &= c
                    conditions.append(comb)

            else:
                parts = [p.strip() for p in val_str.split(',')]
                inc = [p for p in parts if not p.startswith('!')]
                exc = [p.lstrip('!') for p in parts if p.startswith('!')]

                col_conds = []
                if inc:
                    apply_map = (ENABLE_FLO_MAPPING and filter_col == 'FLO')
                    acc = parse_filter_value(','.join(inc), ORDER_MAPPING.get(filter_col), apply_flo_mapping=apply_map)

                    if acc:
                        # SPACE STRIPPER FOR STOCK DATA
                        series_to_check = filtered_df[stock_col].astype(str).str.replace(r'\s+', '', regex=True).str.upper()
                        if apply_map:
                            series_to_check = series_to_check.replace(['VSL', 'SL'], 'FNT')
                        # SPACE STRIPPER FOR FILTER LOGIC
                        acc_clean = [str(x).replace(" ", "").upper() for x in acc]
                        col_conds.append(series_to_check.isin(acc_clean))
                if exc:
                    apply_map = (ENABLE_FLO_MAPPING and filter_col == 'FLO')
                    acc_exc = parse_filter_value(','.join(exc), ORDER_MAPPING.get(filter_col), apply_flo_mapping=apply_map)

                    if acc_exc:
                        # SPACE STRIPPER FOR STOCK DATA
                        series_to_check = filtered_df[stock_col].astype(str).str.replace(r'\s+', '', regex=True).str.upper()
                        if apply_map:
                            series_to_check = series_to_check.replace(['VSL', 'SL'], 'FNT')
                        # SPACE STRIPPER FOR FILTER LOGIC
                        acc_exc_clean = [str(x).replace(" ", "").upper() for x in acc_exc]
                        col_conds.append(~series_to_check.isin(acc_exc_clean))

                if col_conds:
                    comb = pd.Series(True, index=filtered_df.index)
                    for c in col_conds: comb &= c
                    conditions.append(comb)
        except:
            continue

    if not conditions: return pd.DataFrame()
    final_cond = pd.Series(True, index=filtered_df.index)
    for c in conditions: final_cond &= c

    return filtered_df[~final_cond] if remarks == 'EXCLUDE' else filtered_df[final_cond]

def apply_exclusion_rules(df, rules_df):
    if df.empty or rules_df.empty:
        return df

    df_clean = df.copy()

    column_mapping = {
        'SHP': 'Shp', 'QUA': 'Qua', 'CUT': 'Cut', 'POL': 'Pol', 'SYM': 'Sym',
        'FLO': 'Flo', 'LAB': 'Lab', 'DISC%': 'Disc%', 'SHADE': 'Shade', 'MILKY': 'Milky',
        'E.FACET': 'E.Facet', 'TOIN_NAME': 'TOIN_NAME', 'SOIN_NAME': 'SOIN_NAME',
        'GIRDLE_OPEN_NAME': 'GIRDLE_OPEN_NAME', 'PAV_OPEN_NAME': 'PAV_OPEN_NAME',
        'TBINC_NAME': 'TBINC_NAME', 'SBINC_NAME': 'SBINC_NAME', 'MEASUREMENT': 'Measurement',
        'COLOR NAME': 'Color', 'CARAT': 'Carat', 'RESERVE%': 'Reserve%',
        'RATIO': 'Ratio', 'RATE': 'Rate', 'RAPRATE': 'RapRate', 'AMOUNT': 'Amount',
        'TABLE INCL': 'Table Incl', 'TAB': 'Tab', 'C/A': 'C/A', 'DEP': 'Dep', 'C/H': 'C/H',
        'P/A': 'P/A', 'P/H': 'P/H', 'GIRDLE': 'Girdle', 'DIA': 'Dia',
        'SGP COMMENT': 'SGPComment', 'SGPCOMMENT': 'SGPComment',
        'KEYTOSYMBOL': 'KeyToSymbol', 'COMMENT': 'Comment', 'COLOR DESC': 'Color Desc',
        'EYE CLEAN': 'EYE Clean', 'LUSTER': 'Luster', 'CULET': 'Culet'
    }

    def clean_str(val):
        if pd.isna(val) or val == "": return None
        s = str(val).strip()
        if s.endswith('.0'): s = s[:-2]
        elif s.endswith('.00'): s = s[:-3]
        return s

    rows_to_drop_mask = pd.Series(False, index=df_clean.index)
    LONG_TEXT_COLS = ['SGP COMMENT', 'SGPCOMMENT', 'KEYTOSYMBOL', 'COMMENT', 'COLOR DESC']

    for _, rule in rules_df.iterrows():
        rule_mask = pd.Series(True, index=df_clean.index)
        has_criteria = False

        for filter_col, filter_value in rule.items():
            if filter_col in ['SOLD PARTY', 'NORMALIZED PARTY', 'REMARKS', 'SALES_PERSON', 'COUNTRY', 'CITY'] \
               or pd.isna(filter_value) or str(filter_value).strip() == '':
                continue

            stock_col = column_mapping.get(filter_col)
            if not stock_col or stock_col not in df_clean.columns:
                matched_col = next((c for c in df_clean.columns if c.lower() == filter_col.lower()), None)
                if matched_col: stock_col = matched_col
                else: continue

            has_criteria = True
            val_str = str(filter_value).strip().upper()

            try:
                if filter_col in ['CARAT', 'RESERVE%', 'DISC%', 'RATIO', 'RATE', 'RAPRATE',
                                  'AMOUNT', 'TAB', 'C/A', 'DEP', 'C/H', 'P/A', 'P/H', 'GIRDLE', 'DIA']:
                    col_cond = parse_carat_filter(df_clean[stock_col], val_str) if filter_col == 'CARAT' else parse_numeric_filter(df_clean[stock_col], val_str)

                elif filter_col in LONG_TEXT_COLS:
                    is_keep_only = False
                    if val_str.startswith('!'):
                        is_keep_only = True
                        val_str = val_str[1:].strip()

                    vals_to_match = [v.strip() for v in val_str.split(',') if v.strip()]
                    series_to_check = df_clean[stock_col].fillna('').astype(str)
                    pattern = '|'.join([rf"(?<!\w){re.escape(v)}(?!\w)" for v in vals_to_match])
                    match_mask = series_to_check.str.contains(pattern, case=False, regex=True, na=False)

                    if is_keep_only: col_cond = ~match_mask
                    else: col_cond = match_mask

                else:
                    is_keep_only = False
                    if val_str.startswith('!'):
                        is_keep_only = True
                        val_str = val_str[1:].strip()

                    apply_map = (ENABLE_FLO_MAPPING and filter_col == 'FLO')
                    vals_to_match = parse_filter_value(val_str, ORDER_MAPPING.get(filter_col), apply_flo_mapping=apply_map)

                    if vals_to_match:
                        # SPACE STRIPPER FOR STOCK DATA
                        series_to_check = df_clean[stock_col].astype(str).str.replace(r'\s+', '', regex=True).str.upper()
                        if apply_map: series_to_check = series_to_check.replace(['VSL', 'SL'], 'FNT')

                        # SPACE STRIPPER FOR FILTER LOGIC (NO_NEED RULES)
                        clean_vals_to_match = [str(v).replace(" ", "").upper() for v in vals_to_match]

                        if is_keep_only: col_cond = ~series_to_check.isin(clean_vals_to_match)
                        else: col_cond = series_to_check.isin(clean_vals_to_match)
                    else:
                        col_cond = pd.Series(False, index=df_clean.index)

                rule_mask = rule_mask & col_cond
            except:
                pass

        if has_criteria:
            rows_to_drop_mask = rows_to_drop_mask | rule_mask

    df_clean['NoNeed_Fail'] = rows_to_drop_mask

    def extract_pair_id(text):
        if pd.isna(text): return None
        match = re.search(r'(?:PAIRED\s+WITH|MATCH\s+PAIR)\s*[^\d]*([\d\.]+)', str(text), re.IGNORECASE)
        if match: return clean_str(match.group(1))
        return None

    if 'Extracted_Pair_ID' not in df_clean.columns:
        if 'SGPComment' in df_clean.columns:
            df_clean['Extracted_Pair_ID'] = df_clean['SGPComment'].apply(extract_pair_id)
        else:
            df_clean['Extracted_Pair_ID'] = None

    df_clean['Safe_Serial'] = df_clean['Serial No'].apply(clean_str)
    df_clean['Pair_Key'] = df_clean['Extracted_Pair_ID'].apply(clean_str)

    failed_serials = set(df_clean.loc[df_clean['NoNeed_Fail'], 'Safe_Serial'].dropna())
    failed_pair_keys = set(df_clean.loc[df_clean['NoNeed_Fail'], 'Pair_Key'].dropna())
    failed_serials.discard("")
    failed_pair_keys.discard("")

    additional_fails = df_clean['Pair_Key'].isin(failed_serials) | df_clean['Safe_Serial'].isin(failed_pair_keys)
    keep_mask = ~(df_clean['NoNeed_Fail'] | additional_fails)

    result = df_clean[keep_mask].copy()
    drop_cols = ['NoNeed_Fail', 'Safe_Serial', 'Pair_Key']
    result.drop(columns=[c for c in drop_cols if c in result.columns], inplace=True)
    return result

def apply_pair_destruction_sync(df):
    if df.empty or 'SGPComment' not in df.columns:
        return df

    def clean_str(val):
        if pd.isna(val) or val == "": return None
        s = str(val).strip()
        if s.endswith('.0'): s = s[:-2]
        elif s.endswith('.00'): s = s[:-3]
        return s

    def extract_pair_id(text):
        if pd.isna(text): return None
        match = re.search(r'(?:PAIRED\s+WITH|MATCH\s+PAIR)\s*[^\d]*([\d\.]+)', str(text), re.IGNORECASE)
        if match: return clean_str(match.group(1))
        return None

    working_df = df.copy()
    working_df['Safe_Serial'] = working_df['Serial No'].apply(clean_str)
    working_df['Extracted_Pair_ID'] = working_df['SGPComment'].apply(extract_pair_id)

    valid_serials = set(working_df['Safe_Serial'].dropna())
    valid_serials.discard("")

    has_pair_id = working_df['Extracted_Pair_ID'].notna() & (working_df['Extracted_Pair_ID'] != "")
    pair_is_valid = working_df['Extracted_Pair_ID'].isin(valid_serials)

    mask = ~has_pair_id | pair_is_valid
    result_df = working_df[mask].copy()
    result_df.drop(columns=['Safe_Serial', 'Extracted_Pair_ID'], inplace=True, errors='ignore')
    return result_df

def sort_and_format_result(df, salesperson=None, all_stocks_df=None, party_name="Unknown", serial_to_types_map=None):
    if df.empty: return df

    shape_sort_map = {shape: i for i, shape in enumerate(SHAPE_ORDER)}
    df['Shape_Sort_Key'] = df['Shp'].map(shape_sort_map).fillna(999)
    df['Color_Cat'] = pd.Categorical(df['Color'], categories=COLOR_ORDER, ordered=True)
    df['Qua_Cat'] = pd.Categorical(df['Qua'], categories=CLARITY_ORDER, ordered=True)
    df['Carat_Rank'] = df['Carat'].apply(get_carat_rank)

    df.sort_values(by=['Shape_Sort_Key', 'Carat_Rank', 'Color_Cat', 'Qua_Cat'],
                   ascending=[True, True, True, True], inplace=True)

    df.drop(columns=['Carat_Rank', 'Shape_Sort_Key', 'Color_Cat', 'Qua_Cat'], inplace=True, errors='ignore')

    if ENABLE_PAIR_MODE:
        df = apply_pair_destruction_sync(df)

    if ENABLE_TYPE_COLUMNS and serial_to_types_map:
        df['Types_Found_Lookup'] = df['Serial No'].map(serial_to_types_map).fillna("")
        for type_col in STATIC_TYPE_SEQUENCE:
            keyword = type_col.replace("Type_", "")
            df[type_col] = df['Types_Found_Lookup'].apply(
                lambda x: "YES" if keyword in str(x).upper() else ""
            )
        df.drop(columns=['Types_Found_Lookup'], inplace=True)

    if SHOW_CONCAT_STATUS:
        def make_status(row):
            raw_tol = str(row.get('Tolerance_Type', '')).strip().upper()
            if raw_tol == 'NEARBY': tol = "Near by"
            elif raw_tol == 'EXACT': tol = "Exact"
            else: tol = raw_tol.title()

            found_type = ""
            for t_col in STATIC_TYPE_SEQUENCE:
                if str(row.get(t_col, '')).upper() == 'YES':
                    found_type = t_col.replace("Type_", "")
                    break
            return f"{tol}|{found_type}" if found_type else tol

        df['Match_Status_Concatenated'] = df.apply(make_status, axis=1)

    if SHOW_BASIC_SUPERTAG or SHOW_ADVANCED_SUPERTAG:
        if 'Matched_Sales_Tag' not in df.columns:
            df['Matched_Sales_Tag'] = ""

        def parse_tags(tag):
            if pd.isna(tag) or str(tag) == "": return "", ""
            tag = str(tag)
            if '|' in tag:
                tag17 = tag
                tag6 = "-".join(tag.split('|')[:6])
                return tag6, tag17
            else:
                return tag, ""

        tags = df['Matched_Sales_Tag'].apply(parse_tags)
        if SHOW_BASIC_SUPERTAG:
            df['Sales_Supertag_6'] = tags.apply(lambda x: x[0])
        if SHOW_ADVANCED_SUPERTAG:
            df['Sales_Supertag_17'] = tags.apply(lambda x: x[1])

    base_cols_to_drop = ['Super Tag', 'Carat_grp', 'Carat_grp_STD', 'Carat_grp_HK', 'Super Tag STD', 'Super Tag HK', 'Types_Found', 'Safe_Serial', 'Flo_Norm', 'Blind', 'SaleType', 'YourBidPrice', 'BidRate', 'BidAmount', 'Fper', 'Famount']

    if salesperson is not None and not GENERATE_MERGED_MASTER:
        base_cols_to_drop.extend(['Sales_Supertag_6', 'Sales_Supertag_17'])

    clean_cols_to_drop = [c for c in df.columns if c.endswith('_CLEAN')]
    base_cols_to_drop.extend(clean_cols_to_drop)
    if not ENABLE_SALES_SUPERTAG_COLUMN:
        base_cols_to_drop.append('Matched_Sales_Tag')

    df.drop(columns=[c for c in base_cols_to_drop if c in df.columns], inplace=True)
    df.reset_index(drop=True, inplace=True)

    desired_sequence = list(get_desired_sequence_for_salesperson(salesperson))
    current_cols = df.columns.tolist()

    type_cols_present = [c for c in current_cols if c.startswith('Type_')]
    sorted_type_cols = sorted(type_cols_present, key=lambda x: STATIC_TYPE_SEQUENCE.index(x) if x in STATIC_TYPE_SEQUENCE else 999)

    end_cols = sorted_type_cols[:]
    if ENABLE_SALES_SUPERTAG_COLUMN and 'Matched_Sales_Tag' in current_cols:
        end_cols.append('Matched_Sales_Tag')

    if 'Pair_Group_ID' in current_cols: end_cols.append('Pair_Group_ID')
    if 'Is_Pair_Protected' in current_cols: end_cols.append('Is_Pair_Protected')

    if 'Sales_Supertag_6' in current_cols: end_cols.append('Sales_Supertag_6')
    if 'Sales_Supertag_17' in current_cols: end_cols.append('Sales_Supertag_17')

    if SHOW_REF_STONE_ID and 'Ref_Stone_Id' in current_cols: end_cols.append('Ref_Stone_Id')
    if SHOW_CONCAT_STATUS and 'Match_Status_Concatenated' in current_cols: end_cols.append('Match_Status_Concatenated')

    start_cols = [c for c in desired_sequence if c in current_cols]
    exclude_cols = set(start_cols + end_cols)
    middle_cols = [c for c in current_cols if c not in exclude_cols]

    return df[start_cols + middle_cols + end_cols]

def auto_adjust_column_widths(worksheet, df_data=None):
    FIXED_WIDTH_COLS = ["DETAIL", "PHYGITAL", "IMAGE", "VIDEO", "CERTIFICATE", "LAYOUT"]
    CONTENT_TRACK_COLS = ["Sales_Supertag_6", "Sales_Supertag_17", "Matched_Sales_Tag", "Ref_Stone_Id", "Match_Status_Concatenated"]

    # --- FAST PATH: For giant data sheets (uses Pandas for instant width calculation) ---
    if df_data is not None and not df_data.empty:
        for i, column_cell in enumerate(worksheet[1], 1):
            col_letter = get_column_letter(i)
            col_name = str(column_cell.value).strip() if column_cell.value else ""

            if col_name in FIXED_WIDTH_COLS:
                worksheet.column_dimensions[col_letter].width = 12
                continue

            max_len = len(col_name) + 2

            if col_name in df_data.columns:
                try:
                    series_len = df_data[col_name].astype(str).map(len).max()
                    if pd.notna(series_len):
                        max_len = max(max_len, series_len)
                except:
                    pass

            is_tag_col = col_name in CONTENT_TRACK_COLS
            limit = 100 if is_tag_col else 40
            adjusted_width = min((max_len + 2) * 1.1, limit)
            worksheet.column_dimensions[col_letter].width = adjusted_width

    # --- SAFE PATH: For the Navigation sheet (calculates standard Excel cells) ---
    else:
        for col in worksheet.columns:
            max_length = 0
            col_letter = col[0].column_letter
            col_name = str(col[0].value).strip() if col[0].value else ""

            if col_name in FIXED_WIDTH_COLS:
                worksheet.column_dimensions[col_letter].width = 12
                continue

            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass

            is_tag_col = col_name in CONTENT_TRACK_COLS
            limit = 100 if is_tag_col else 40
            adjusted_width = min((max_length + 2) * 1.1, limit)
            worksheet.column_dimensions[col_letter].width = adjusted_width
        return

def apply_styles_to_worksheet(worksheet):
    thin_side = Side(style='thin')
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=False)
    header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")

    filter_match_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    search_match_fill = PatternFill(start_color="E0F7FA", end_color="E0F7FA", fill_type="solid")

    if ENABLE_NEW_FILL_LOGIC:
        exact_match_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
        nearby_match_fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
        maybe_match_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    else:
        exact_match_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
        nearby_match_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        maybe_match_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    headers = [cell.value for cell in worksheet[1]]

    try: is_nearby_idx = headers.index('Is_Nearby')
    except: is_nearby_idx = -1

    try: is_filter_match_idx = headers.index('Is_Filter_Match')
    except: is_filter_match_idx = -1

    try: match_type_idx = headers.index('Match_Type')
    except: match_type_idx = -1

    try: tolerance_type_idx = headers.index('Tolerance_Type')
    except: tolerance_type_idx = -1

    header_font = Font(bold=True)
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border
        cell.alignment = align_center

    if worksheet.title == "Navigation":
        header_font_nav = Font(bold=True, color="0000FF")
        for cell in worksheet[1]: cell.font = header_font_nav
        for row in worksheet.iter_rows():
            for cell in row:
                cell.border = thin_border
                cell.alignment = align_center
        return

    rows = list(worksheet.iter_rows(min_row=2))

    for row_cells in rows:
        fill_to_apply = None

        if is_nearby_idx != -1 and row_cells[is_nearby_idx].value is False:
             fill_to_apply = exact_match_fill
        elif is_filter_match_idx != -1 and row_cells[is_filter_match_idx].value is True:
             fill_to_apply = filter_match_fill
        elif match_type_idx != -1 and row_cells[match_type_idx].value == 'SEARCH':
             fill_to_apply = search_match_fill
        elif tolerance_type_idx != -1:
             tol_val = row_cells[tolerance_type_idx].value
             if tol_val == 'NEARBY': fill_to_apply = nearby_match_fill
             elif tol_val == 'MAYBE': fill_to_apply = maybe_match_fill
        elif is_nearby_idx != -1 and row_cells[is_nearby_idx].value is True and not ENABLE_NEW_FILL_LOGIC:
             fill_to_apply = exact_match_fill

        for cell in row_cells:
            cell.border = thin_border
            cell.alignment = align_center
            if fill_to_apply:
                cell.fill = fill_to_apply

    blue_underline = Font(color="0000FF", underline="single")

    for col_idx, header in enumerate(headers):
        if header in DECIMAL_COLUMNS:
            col_letter = get_column_letter(col_idx + 1)
            for cell in worksheet[col_letter][1:]:
                 if isinstance(cell.value, (int, float)):
                     cell.number_format = numbers.FORMAT_NUMBER_00

        if header == 'Sr.No':
            col_letter = get_column_letter(col_idx + 1)
            for cell in worksheet[col_letter][1:]:
                cell.font = blue_underline

def process_party_from_profile(
    party_profile_rows: pd.DataFrame,
    stocks_df: pd.DataFrame,
    top_n: int,
    all_change_format_filters: pd.DataFrame,
    default_tolerance_level: int,
    tolerance_df: pd.DataFrame,
    current_salesperson: str,
    stocks_by_tag: dict,
    stocks_by_key: dict,
    stocks_by_shape: dict
):
    norm_party_name = party_profile_rows['Normalized Party'].iloc[0]
    party_display_name = party_profile_rows['SoldParty'].iloc[0] if 'SoldParty' in party_profile_rows.columns else "Unknown Party"
    country = party_profile_rows['Country'].iloc[0] if 'Country' in party_profile_rows.columns else ''
    city = party_profile_rows['City'].iloc[0] if 'City' in party_profile_rows.columns else ''

    exact_match_serials = set()
    serial_to_types_map = {}

    full_profile_tags = party_profile_rows[['Super Tag', 'Types_Found']].dropna(subset=['Super Tag'])

    active_match_columns = ['Shp', 'Carat_grp', 'Lab']
    if INCLUDE_COLOR: active_match_columns.append('Color')
    if INCLUDE_CLARITY: active_match_columns.append('Qua')
    if INCLUDE_FLORO: active_match_columns.append('Flo')
    if INCLUDE_CUT: active_match_columns.append('Cut')
    if INCLUDE_POLISH: active_match_columns.append('Pol')
    if INCLUDE_SYMMETRY: active_match_columns.append('Sym')
    if INCLUDE_SHADE: active_match_columns.append('Shade')
    if INCLUDE_MILKY: active_match_columns.append('Milky')
    if INCLUDE_TABLE_BLACK: active_match_columns.append('TBINC_NAME')
    if INCLUDE_SIDE_BLACK: active_match_columns.append('SBINC_NAME')
    if INCLUDE_TABLE_OPEN: active_match_columns.append('TOIN_NAME')
    if INCLUDE_SIDE_OPEN: active_match_columns.append('SOIN_NAME')
    if INCLUDE_GIRDLE_OPEN: active_match_columns.append('GIRDLE_OPEN_NAME')
    if INCLUDE_PAVILION_OPEN: active_match_columns.append('PAV_OPEN_NAME')

    for _, row in full_profile_tags.iterrows():
        tag = row['Super Tag']
        if tag == 'DUMMY_TAG': continue
        types_str = str(row['Types_Found'])

        matches = []
        if '|' in tag:
            tag_parts = tag.split('|')
            shape_key = tag_parts[0] if tag_parts else None

            subset_df = stocks_df
            if shape_key and shape_key in stocks_by_shape:
                 subset_df = stocks_by_shape[shape_key]

            if len(tag_parts) in (len(FULL_TAG_MAP), 6):
                criteria = dict(zip(FULL_TAG_MAP if len(tag_parts) == len(FULL_TAG_MAP) else BASIC_TAG_MAP, tag_parts))
                mask = pd.Series(True, index=subset_df.index)

                for col in active_match_columns:
                    if col in criteria:
                        val = criteria[col]
                        if val in ['ANY', '*']: continue

                        if col == 'Carat_grp':
                            mask &= ((subset_df['Carat_grp_STD'] == val) | (subset_df['Carat_grp_HK'] == val))
                        elif col == 'Flo':
                            val_norm = 'FNT' if (ENABLE_FLO_MAPPING and val in ['VSL', 'SL']) else val
                            if ENABLE_FLO_MAPPING:
                                stock_val_norm = subset_df['Flo'].replace(['VSL', 'SL'], 'FNT')
                                mask &= (stock_val_norm == val_norm)
                            else:
                                mask &= (subset_df['Flo'] == val)
                        elif f'{col}_CLEAN' in subset_df.columns:
                            val_clean = str(val).replace(' ', '').upper()
                            mask &= (subset_df[f'{col}_CLEAN'] == val_clean)
                        elif col in subset_df.columns:
                            val_clean = str(val).replace(' ', '').upper()
                            stock_val_norm = subset_df[col].astype(str).str.replace(r'\s+', '', regex=True).str.upper()
                            mask &= (stock_val_norm == val_clean)

                matches = subset_df.loc[mask, 'Serial No'].tolist()

        else:
            if tag in stocks_by_tag:
                matches = stocks_by_tag[tag]['Serial No'].tolist()

        if matches:
            exact_match_serials.update(matches)
            for s in matches:
                if s in serial_to_types_map:
                    serial_to_types_map[s] = serial_to_types_map[s] + "," + types_str
                else:
                    serial_to_types_map[s] = types_str

    party_filter_rules_df = pd.DataFrame()
    if not all_change_format_filters.empty and 'NORMALIZED PARTY' in all_change_format_filters.columns:

        # 1. Clean the salesperson names (removes dots, strips extra spaces, makes uppercase)
        clean_filter_sp = all_change_format_filters['SALES_PERSON'].astype(str).str.replace('.', '', regex=False).str.strip().str.upper()
        clean_current_sp = str(current_salesperson).replace('.', '').strip().upper()

        # 2. Match using the perfectly cleaned strings
        party_filter_rules_df = all_change_format_filters[
            (all_change_format_filters['NORMALIZED PARTY'] == norm_party_name) &
            (clean_filter_sp == clean_current_sp)
        ]

    current_tolerance_level = default_tolerance_level
    enable_carat_tolerance = False
    if not tolerance_df.empty and current_salesperson:
        tol_match = tolerance_df[
            (tolerance_df['NORMALIZED PARTY'] == norm_party_name) &
            (tolerance_df['SALES_PERSON'] == current_salesperson)
        ]
        if not tol_match.empty:
            rule = tol_match.iloc[0]
            t2 = str(rule.get('TOLERANCE +-2', '')).strip().upper()
            t1 = str(rule.get('TOLERANCE +-1', '')).strip().upper()
            if t2 == 'YES': current_tolerance_level = 2
            elif t1 == 'YES': current_tolerance_level = 1
            elif t2 == 'NO' and t1 == 'NO': current_tolerance_level = 0
            carat_tol = str(rule.get('CARAT_GRP_TOLERANCE', '')).strip().upper()
            if carat_tol == 'YES': enable_carat_tolerance = True

    do_exact = True
    do_tolerance = (current_tolerance_level > 0)

    if not party_filter_rules_df.empty and 'PAST_SALES' in party_filter_rules_df.columns:
        past_val = str(party_filter_rules_df.iloc[0]['PAST_SALES']).strip().upper()
        if past_val == 'NO': do_exact, do_tolerance = False, False
        elif past_val == 'YES': do_exact, do_tolerance = True, False

    final_result_list = []
    seen_serials = set()

    if not party_filter_rules_df.empty:
        for idx, party_filter_rule in party_filter_rules_df.iterrows():
            filter_matches_df = apply_advanced_filters(stocks_df, party_filter_rule)
            if not filter_matches_df.empty:
                filter_matches_df = filter_matches_df[~filter_matches_df['Serial No'].isin(seen_serials)].copy()
                if not filter_matches_df.empty:
                    filter_matches_df['Is_Filter_Match'] = True
                    filter_matches_df['Match_Type'] = 'FILTER'
                    filter_matches_df['Tolerance_Type'] = 'NONE'
                    final_result_list.append(filter_matches_df)
                    seen_serials.update(filter_matches_df['Serial No'].tolist())

    tag_rows = party_profile_rows.sort_values('Count', ascending=False)
    if top_n is not None:
        tag_rows = tag_rows.head(top_n)

    top_n_tags = [t for t in tag_rows['Super Tag'].tolist() if t != 'DUMMY_TAG']
    tag_match_logic = dict(zip(tag_rows['Super Tag'], tag_rows['Match_Logic_Type']))
    tag_types_map = dict(zip(tag_rows['Super Tag'], tag_rows['Types_Found']))

    tag_ref_stone_map = {}
    if SHOW_REF_STONE_ID and 'Ref_Stone_Id' in tag_rows.columns:
        tag_ref_stone_map = dict(zip(tag_rows['Super Tag'], tag_rows['Ref_Stone_Id']))

    if do_exact and top_n_tags:
        for tag in top_n_tags:
            matches_df = pd.DataFrame()
            if '|' in tag:
                tag_parts = tag.split('|')
                shape_key = tag_parts[0] if tag_parts else None
                subset_df = stocks_df
                if shape_key and shape_key in stocks_by_shape:
                    subset_df = stocks_by_shape[shape_key]

                if len(tag_parts) in (len(FULL_TAG_MAP), 6):
                    criteria = dict(zip(FULL_TAG_MAP if len(tag_parts) == len(FULL_TAG_MAP) else BASIC_TAG_MAP, tag_parts))
                    mask = pd.Series(True, index=subset_df.index)

                    for col in active_match_columns:
                        if col in criteria:
                            val = criteria[col]
                            if val in ['ANY', '*']: continue

                            if col == 'Carat_grp':
                                 mask &= ((subset_df['Carat_grp_STD'] == val) | (subset_df['Carat_grp_HK'] == val))
                            elif col == 'Flo':
                                val_norm = 'FNT' if (ENABLE_FLO_MAPPING and val in ['VSL', 'SL']) else val
                                if ENABLE_FLO_MAPPING:
                                    stock_val_norm = subset_df['Flo'].replace(['VSL', 'SL'], 'FNT')
                                    mask &= (stock_val_norm == val_norm)
                                else:
                                    mask &= (subset_df['Flo'] == val)
                            elif f'{col}_CLEAN' in subset_df.columns:
                                val_clean = str(val).replace(' ', '').upper()
                                mask &= (subset_df[f'{col}_CLEAN'] == val_clean)
                            elif col in subset_df.columns:
                                val_clean = str(val).replace(' ', '').upper()
                                stock_val_norm = subset_df[col].astype(str).str.replace(r'\s+', '', regex=True).str.upper()
                                mask &= (stock_val_norm == val_clean)

                    matches_df = subset_df[mask].copy()

            else:
                if tag in stocks_by_tag:
                    matches_df = stocks_by_tag[tag].copy()

            if not matches_df.empty:
                matches_df = matches_df[~matches_df['Serial No'].isin(seen_serials)]
                if not matches_df.empty:
                    matches_df.drop_duplicates(subset=['Serial No'], keep='first', inplace=True)
                    matches_df['Is_Nearby'] = False
                    matches_df['Match_Type'] = tag_match_logic.get(tag, 'STANDARD')
                    matches_df['Matched_Sales_Tag'] = tag
                    if SHOW_REF_STONE_ID:
                        matches_df['Ref_Stone_Id'] = tag_ref_stone_map.get(tag, "")
                    matches_df['Tolerance_Type'] = 'EXACT'
                    seen_serials.update(matches_df['Serial No'].tolist())
                    final_result_list.append(matches_df)

    if do_tolerance and top_n_tags:
        for tag in top_n_tags:
            if tag_match_logic.get(tag) == 'SEARCH': continue

            if '|' in tag:
                tag_parts = tag.split('|')
                shape_key = tag_parts[0] if tag_parts else None
                subset_df = stocks_df
                if shape_key and shape_key in stocks_by_shape:
                    subset_df = stocks_by_shape[shape_key]

                criteria = {}
                if len(tag_parts) in (len(FULL_TAG_MAP), 6):
                    criteria = dict(zip(FULL_TAG_MAP if len(tag_parts) == len(FULL_TAG_MAP) else BASIC_TAG_MAP, tag_parts))

                if criteria:
                    mask = pd.Series(True, index=subset_df.index)

                    c_val = criteria.get('Carat_grp')
                    if c_val and c_val not in ['ANY', '*'] and str(c_val).strip() != "":
                        if enable_carat_tolerance:
                            neighbors = get_neighbor_carat_grps(c_val)
                            mask &= ((subset_df['Carat_grp_STD'].isin(neighbors)) | (subset_df['Carat_grp_HK'].isin(neighbors)))
                        else:
                            mask &= ((subset_df['Carat_grp_STD'] == c_val) | (subset_df['Carat_grp_HK'] == c_val))

                    strict_cols = ['Shp', 'Lab']
                    relax_cols = {}

                    def assign_col(flag_include, flag_tol, col_name, order_list=None):
                        if flag_include:
                            if flag_tol and order_list:
                                relax_cols[col_name] = order_list
                            else:
                                 strict_cols.append(col_name)

                    assign_col(INCLUDE_COLOR, ENABLE_COLOR_TOLERANCE, 'Color', COLOR_ORDER)
                    assign_col(INCLUDE_CLARITY, ENABLE_CLARITY_TOLERANCE, 'Qua', CLARITY_ORDER)
                    assign_col(INCLUDE_FLORO, ENABLE_FLORO_TOLERANCE, 'Flo', FLUORESCENCE_ORDER)

                    assign_col(INCLUDE_CUT, ENABLE_CUT_TOLERANCE, 'Cut', CUT_POL_SYM_ORDER)
                    assign_col(INCLUDE_POLISH, ENABLE_POLISH_TOLERANCE, 'Pol', CUT_POL_SYM_ORDER)
                    assign_col(INCLUDE_SYMMETRY, ENABLE_SYMMETRY_TOLERANCE, 'Sym', CUT_POL_SYM_ORDER)

                    assign_col(INCLUDE_SHADE, ENABLE_SHADE_TOLERANCE, 'Shade', SHADE_ORDER)
                    assign_col(INCLUDE_MILKY, ENABLE_MILKY_TOLERANCE, 'Milky', MILKY_ORDER)

                    assign_col(INCLUDE_TABLE_BLACK, ENABLE_TABLE_BLACK_TOLERANCE, 'TBINC_NAME', TBINC_SBINC_ORDER)
                    assign_col(INCLUDE_SIDE_BLACK, ENABLE_SIDE_BLACK_TOLERANCE, 'SBINC_NAME', TBINC_SBINC_ORDER)

                    assign_col(INCLUDE_TABLE_OPEN, ENABLE_TABLE_OPEN_TOLERANCE, 'TOIN_NAME', TOIN_SOIN_GIRDLE_PAV_OPEN_ORDER)
                    assign_col(INCLUDE_SIDE_OPEN, ENABLE_SIDE_OPEN_TOLERANCE, 'SOIN_NAME', TOIN_SOIN_GIRDLE_PAV_OPEN_ORDER)
                    assign_col(INCLUDE_GIRDLE_OPEN, ENABLE_GIRDLE_OPEN_TOLERANCE, 'GIRDLE_OPEN_NAME', TOIN_SOIN_GIRDLE_PAV_OPEN_ORDER)
                    assign_col(INCLUDE_PAVILION_OPEN, ENABLE_PAVILION_OPEN_TOLERANCE, 'PAV_OPEN_NAME', TOIN_SOIN_GIRDLE_PAV_OPEN_ORDER)

                    for col in strict_cols:
                        if col in criteria:
                            val = criteria[col]
                            if val and val not in ['ANY', '*'] and str(val).strip() != "":
                                if f'{col}_CLEAN' in subset_df.columns:
                                    val_clean = str(val).replace(' ', '').upper()
                                    mask &= (subset_df[f'{col}_CLEAN'] == val_clean)
                                elif col in subset_df.columns:
                                    val_clean = str(val).replace(' ', '').upper()
                                    stock_val_norm = subset_df[col].astype(str).str.replace(r'\s+', '', regex=True).str.upper()
                                    mask &= (stock_val_norm == val_clean)

                    for col, order in relax_cols.items():
                        if col in criteria:
                            val = criteria[col]
                            if val and val not in ['ANY', '*'] and str(val).strip() != "":
                                val_lookup = str(val).replace(' ', '').upper()
                                if col == 'Flo' and ENABLE_FLO_MAPPING and val in ['VSL', 'SL']:
                                    val_lookup = 'FNT'

                                tol_vals = get_tolerance_values(val_lookup, order, current_tolerance_level)

                                if col == 'Flo' and ENABLE_FLO_MAPPING:
                                    stock_val_norm = subset_df['Flo'].replace(['VSL', 'SL'], 'FNT')
                                    mask &= (stock_val_norm.isin(tol_vals))
                                elif f'{col}_CLEAN' in subset_df.columns:
                                    mask &= (subset_df[f'{col}_CLEAN'].isin(tol_vals))
                                elif col in subset_df.columns:
                                    stock_val_norm = subset_df[col].astype(str).str.replace(r'\s+', '', regex=True).str.upper()
                                    mask &= (stock_val_norm.isin(tol_vals))

                    subset = subset_df[mask].copy()
                    subset = subset[~subset['Serial No'].isin(seen_serials)]

                    if not subset.empty:
                        score = 0
                        max_score = 0

                        if INCLUDE_COLOR and criteria.get('Color') and criteria.get('Color') not in ['ANY', '*']:
                            score += (subset['Color'] == criteria.get('Color')).astype(int)
                            max_score += 1

                        if INCLUDE_CLARITY and criteria.get('Qua') and criteria.get('Qua') not in ['ANY', '*']:
                            score += (subset['Qua'] == criteria.get('Qua')).astype(int)
                            max_score += 1

                        if INCLUDE_FLORO and criteria.get('Flo') and criteria.get('Flo') not in ['ANY', '*']:
                            target_flo = criteria.get('Flo')
                            if ENABLE_FLO_MAPPING and target_flo in ['VSL', 'SL']:
                                target_flo = 'FNT'

                            stock_flo_check = subset['Flo']
                            if ENABLE_FLO_MAPPING:
                                stock_flo_check = stock_flo_check.replace(['VSL', 'SL'], 'FNT')

                            score += (stock_flo_check == target_flo).astype(int)
                            max_score += 1

                        subset['Tolerance_Type'] = 'MAYBE'
                        threshold = max(1, max_score - 1)
                        subset.loc[score >= threshold, 'Tolerance_Type'] = 'NEARBY'

                        subset['Is_Nearby'] = True
                        subset['Match_Type'] = 'TOLERANCE'
                        subset['Matched_Sales_Tag'] = tag
                        if SHOW_REF_STONE_ID:
                            subset['Ref_Stone_Id'] = tag_ref_stone_map.get(tag, "")

                        current_serials = subset['Serial No'].tolist()
                        seen_serials.update(current_serials)
                        current_types = tag_types_map.get(tag, "")
                        for s in current_serials:
                            if s not in serial_to_types_map:
                                 serial_to_types_map[s] = current_types

                        final_result_list.append(subset)

            else:
                tag_parts = tag.split('-')
                if len(tag_parts) < 6: continue
                shp = tag_parts[0]
                carat_grp = '-'.join(tag_parts[1:-4])
                col, qua, flo, lab = tag_parts[-4], tag_parts[-3], tag_parts[-2], tag_parts[-1]

                carat_grps_to_check = [carat_grp]
                if enable_carat_tolerance:
                    carat_grps_to_check = get_neighbor_carat_grps(carat_grp)

                subset = pd.DataFrame()
                if stocks_by_key is not None:
                    frames = []
                    for cg in carat_grps_to_check:
                        key = (shp, cg, lab)
                        if key in stocks_by_key:
                            frames.append(stocks_by_key[key])
                    if frames:
                        subset = pd.concat(frames, ignore_index=False)

                if subset.empty: continue

                mask = pd.Series(True, index=subset.index)

                if INCLUDE_COLOR:
                    color_range = get_tolerance_values(col, COLOR_ORDER, current_tolerance_level) if ENABLE_COLOR_TOLERANCE else [col]
                    mask &= subset['Color'].isin(color_range)

                if INCLUDE_CLARITY:
                    clarity_range = get_tolerance_values(qua, CLARITY_ORDER, current_tolerance_level) if ENABLE_CLARITY_TOLERANCE else [qua]
                    mask &= subset['Qua'].isin(clarity_range)

                if INCLUDE_FLORO:
                    flo_lookup = 'FNT' if (ENABLE_FLO_MAPPING and flo in ['VSL', 'SL']) else flo
                    fluorescence_range = get_tolerance_values(flo_lookup, FLUORESCENCE_ORDER, current_tolerance_level) if ENABLE_FLORO_TOLERANCE else [flo_lookup]

                    subset_flo_series = subset['Flo'].astype(str).str.upper()
                    if ENABLE_FLO_MAPPING:
                         subset_flo_series = subset_flo_series.replace(['VSL', 'SL'], 'FNT')
                    mask &= subset_flo_series.isin(fluorescence_range)

                tolerance_matches = subset[mask & (~subset['Serial No'].isin(seen_serials))].copy()

                if not tolerance_matches.empty:
                    score_series = pd.Series(0, index=tolerance_matches.index)
                    max_score = 0

                    if INCLUDE_COLOR:
                        score_series += (tolerance_matches['Color'] == col).astype(int)
                        max_score += 1

                    if INCLUDE_CLARITY:
                        score_series += (tolerance_matches['Qua'] == qua).astype(int)
                        max_score += 1

                    if INCLUDE_FLORO:
                        stock_flo_norm = tolerance_matches['Flo'].replace(['VSL', 'SL'], 'FNT') if ENABLE_FLO_MAPPING else tolerance_matches['Flo']
                        target_flo_norm = 'FNT' if (ENABLE_FLO_MAPPING and flo in ['VSL', 'SL']) else flo
                        score_series += (stock_flo_norm == target_flo_norm).astype(int)
                        max_score += 1

                    tolerance_matches['Tolerance_Type'] = 'MAYBE'
                    threshold = max(1, max_score - 1)
                    tolerance_matches.loc[score_series >= threshold, 'Tolerance_Type'] = 'NEARBY'

                    tolerance_matches['Is_Nearby'] = True
                    tolerance_matches['Match_Type'] = 'TOLERANCE'
                    tolerance_matches['Matched_Sales_Tag'] = tag
                    if SHOW_REF_STONE_ID:
                        tolerance_matches['Ref_Stone_Id'] = tag_ref_stone_map.get(tag, "")

                    current_serials = tolerance_matches['Serial No'].tolist()
                    seen_serials.update(current_serials)
                    current_types = tag_types_map.get(tag, "")
                    for s in current_serials:
                        if s not in serial_to_types_map:
                            serial_to_types_map[s] = current_types

                    final_result_list.append(tolerance_matches)

    combined_result = pd.concat(final_result_list) if final_result_list else pd.DataFrame()
    if not combined_result.empty:
        combined_result.drop_duplicates(subset=['Serial No'], keep='first', inplace=True)

    should_apply_filter = False

    if MATCH_STATUS_FILTER_MODE == "ALL":
        should_apply_filter = True
    elif MATCH_STATUS_FILTER_MODE == "NONE":
        should_apply_filter = False
    elif MATCH_STATUS_FILTER_MODE == "SPECIFIC":
        normalized_targets = [normalize_party_name(p) for p in MATCH_STATUS_SPECIFIC_PARTIES]
        if norm_party_name in normalized_targets:
            should_apply_filter = True
        else:
            should_apply_filter = False

    if should_apply_filter and not combined_result.empty and MATCH_STATUS_ALLOWED_TYPES:
        combined_result['temp_status'] = combined_result['Tolerance_Type'].map({
            'EXACT': 'Exact',
            'NEARBY': 'Near by',
            'MAYBE': 'May be',
            'NONE': 'Filter'
        })
        combined_result = combined_result[combined_result['temp_status'].isin(MATCH_STATUS_ALLOWED_TYPES)]
        combined_result.drop(columns=['temp_status'], inplace=True)

    combined_result = sort_and_format_result(
        combined_result,
        salesperson=current_salesperson,
        all_stocks_df=stocks_df,
        party_name=party_display_name,
        serial_to_types_map=serial_to_types_map
    )

    return combined_result, party_display_name, country, city

def format_native_hyperlinks(worksheet, dataframe):
    """Bypasses Excel's 255-char formula limit by applying native XML hyperlinks."""
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

def load_combined_party_profile(perm_path, temp_path):
    def safe_read(path):
        if os.path.exists(path): return pd.read_parquet(path)
        return pd.DataFrame()

    perm = safe_read(perm_path)
    temp = safe_read(temp_path)
    if perm.empty and temp.empty:
        return pd.DataFrame()

    required_cols = [
        'Salesperson', 'Sold Party Original', 'Normalized Party', 'Super Tag', 'Count', 'Match_Logic_Type',
        'SoldParty', 'Country', 'City',
        'BidSaleCount', 'MTD_Label', 'Types_Found', 'Stone_Id'
    ]

    for df in [perm, temp]:
        if not df.empty:
            if 'Sold Party Original' in df.columns and 'Normalized Party' not in df.columns:
                df['Normalized Party'] = df['Sold Party Original'].apply(normalize_party_name)
            if 'Sold Party Original' in df.columns and 'SoldParty' not in df.columns:
                df['SoldParty'] = df['Sold Party Original']

            for col in required_cols:
                if col not in df.columns:
                    if col in ['Count', 'BidSaleCount']:
                        df[col] = 0
                    else:
                        df[col] = ""

    combined = pd.concat([perm, temp], ignore_index=True)

    def agg_match_logic(series):
        return 'STANDARD' if (series == 'STANDARD').any() else 'SEARCH'

    def agg_types_found(series):
        valid = [str(x) for x in series if pd.notna(x) and x != ""]
        return ",".join(list(set(",".join(valid).split(","))))

    tag_stats = (
        combined
        .groupby(['Salesperson', 'Normalized Party', 'Super Tag'])
        .agg(
            Count=('Count', 'sum'),
            BidSaleCount=('BidSaleCount', 'sum'),
            Match_Logic_Type=('Match_Logic_Type', agg_match_logic),
            Types_Found=('Types_Found', agg_types_found),
            Ref_Stone_Id=('Stone_Id', lambda x: x.iloc[0] if not x.empty else '')
        )
        .reset_index()
    )

    meta = (
        combined
        .groupby(['Salesperson', 'Normalized Party'])
        .agg(
            SoldParty=('SoldParty', lambda s: s.dropna().iloc[0] if s.dropna().size else ''),
            Country=('Country', lambda s: s.dropna().iloc[0] if s.dropna().size else ''),
            City=('City', lambda s: s.dropna().iloc[0] if s.dropna().size else ''),
            MTD_Label=('MTD_Label', lambda s: s.dropna().iloc[0] if s.dropna().size else '')
        )
        .reset_index()
    )

    return tag_stats.merge(meta, on=['Salesperson', 'Normalized Party'], how='left')

def generate_salesperson_reports_from_profile(stock_file, party_profile_perm_path, party_profile_temp_path, output_zip_path, filter_file_path, clean_file_path, sales_history_path, top_n, tolerance_level):
    print("Loading Party Profiles...")
    party_profile = load_combined_party_profile(party_profile_perm_path, party_profile_temp_path)

    sales_info_for_sorting = None
    if ENABLE_ADVANCED_SORTING:
        try:
            sales_info_for_sorting = load_sales_history_for_sorting(sales_history_path)
        except Exception as e:
            print(f"Error loading sales history for sorting: {e}")
            sales_info_for_sorting = None

    print("Loading Stock Data...")
    stocks_df = pd.read_excel(stock_file, sheet_name='Sheet1')

    print("Optimizing Stock Data (Vectorized)...")
    str_cols = ['Shp', 'Color', 'Qua', 'Flo', 'Lab', 'Cut', 'Pol', 'Sym', 'Shade', 'Milky']
    for col in str_cols:
        if col in stocks_df.columns:
            stocks_df[col] = stocks_df[col].fillna("").astype(str).str.strip().str.upper()
            stocks_df[col] = stocks_df[col].replace('NAN', '')
            stocks_df[f'{col}_CLEAN'] = stocks_df[col].str.replace(r'\s+', '', regex=True)

    if 'Serial No' in stocks_df.columns: stocks_df.drop_duplicates(subset=['Serial No'], keep='first', inplace=True)

    reverse_stock_map = {}
    for final_name, variations in STOCK_COLUMN_MAPPING.items():
        for var in variations: reverse_stock_map[var.upper()] = final_name
    stocks_df.columns = [reverse_stock_map.get(col.strip().upper(), col.strip()) for col in stocks_df.columns]

    numeric_cols = ['Carat', 'Reserve%', 'Disc%', 'Ratio', 'Rate', 'RapRate', 'Amount', 'Tab', 'C/A', 'Dep', 'C/H', 'P/A', 'P/H', 'Girdle', 'Dia']
    for col in numeric_cols:
        if col in stocks_df.columns: stocks_df[col] = pd.to_numeric(stocks_df[col], errors='coerce')

    stocks_df['Carat_grp_STD'] = stocks_df['Carat'].apply(lambda x: get_carat_range_string(x, country="USA"))
    stocks_df['Carat_grp_HK'] = stocks_df['Carat'].apply(lambda x: get_carat_range_string(x, country="HONGKONG"))

    required_super_tag_cols_base = ['Shp', 'Color', 'Qua', 'Flo', 'Lab']
    if not all(col in stocks_df.columns for col in required_super_tag_cols_base + ['Carat']):
        print("Missing required columns in stock file.")
        return

    print("Building stock indexes (including Shape Optimization)...")
    stocks_by_tag = {}
    stocks_by_key = {}
    stocks_by_shape = {}

    stocks_df['Super Tag STD'] = (stocks_df['Shp'] + '-' + stocks_df['Carat_grp_STD'] + '-' + stocks_df['Color'] + '-' + stocks_df['Qua'] + '-' + stocks_df['Flo'] + '-' + stocks_df['Lab'])
    stocks_df['Super Tag HK'] = (stocks_df['Shp'] + '-' + stocks_df['Carat_grp_HK'] + '-' + stocks_df['Color'] + '-' + stocks_df['Qua'] + '-' + stocks_df['Flo'] + '-' + stocks_df['Lab'])

    for tag, subdf in stocks_df.groupby('Super Tag STD'):
        if tag in stocks_by_tag: stocks_by_tag[tag] = pd.concat([stocks_by_tag[tag], subdf])
        else: stocks_by_tag[tag] = subdf
    for tag, subdf in stocks_df.groupby('Super Tag HK'):
        if tag not in stocks_by_tag: stocks_by_tag[tag] = subdf

    for (shp, grp, lab), subdf in stocks_df.groupby(['Shp', 'Carat_grp_STD', 'Lab']):
        stocks_by_key[(shp, grp, lab)] = subdf
    for (shp, grp, lab), subdf in stocks_df.groupby(['Shp', 'Carat_grp_HK', 'Lab']):
         if (shp, grp, lab) not in stocks_by_key: stocks_by_key[(shp, grp, lab)] = subdf

    for shp, subdf in stocks_df.groupby('Shp'):
        stocks_by_shape[shp] = subdf

    removed_parties, all_change_format_filters, tolerance_df = load_filter_data(filter_file_path)

    exclusion_data = pd.DataFrame()
    if ENABLE_CLEAN_FILE:
        exclusion_data = load_exclusion_data(clean_file_path)

    if removed_parties: party_profile = party_profile[~party_profile['Normalized Party'].isin(removed_parties)]
    salespersons_from_profile = party_profile['Salesperson'].unique().tolist() if not party_profile.empty else []
    salespersons_from_filters = all_change_format_filters['SALES_PERSON'].unique().tolist() if 'SALES_PERSON' in all_change_format_filters.columns else []
    all_salespersons = sorted(list(set(salespersons_from_profile + salespersons_from_filters)))

    if TEST_MODE:
        all_salespersons = [sp for sp in all_salespersons if sp in TEST_SALESPEOPLE]

    normalized_test_parties = []
    if TEST_MODE_PARTY:
        normalized_test_parties = [normalize_party_name(p) for p in TEST_PARTIES]

    print(f"Processing {len(all_salespersons)} salespersons...")

    all_consolidated_data = []

    with zipfile.ZipFile(output_zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for salesperson in all_salespersons:
            print(f"Processing: {salesperson}")
            salesperson_party_reports = {}
            sp_profile = party_profile[party_profile['Salesperson'] == salesperson]

            top_customers_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

            top_customers_parties = set()
            if ENABLE_TOP_CUSTOMERS_HIGHLIGHT:
                if not sp_profile.empty and 'BidSaleCount' in sp_profile.columns:
                    party_scores = (
                        sp_profile
                        .groupby('Normalized Party')['BidSaleCount']
                        .sum()
                        .sort_values(ascending=False)
                    )
                    top_customers_parties = set(party_scores.head(int(TOP_CUSTOMERS_N)).index.tolist())

            global_exclusion_rules = pd.DataFrame()
            party_specific_exclusion = pd.DataFrame()

            if ENABLE_CLEAN_FILE and not exclusion_data.empty:
                sp_exclusion = exclusion_data[exclusion_data['SALES_PERSON'] == salesperson].copy()
                if not sp_exclusion.empty:
                    global_mask = (sp_exclusion['SOLD PARTY'] == 'ALL') & (sp_exclusion['REMARKS'] == 'ALL')
                    global_exclusion_rules = sp_exclusion[global_mask]
                    party_specific_exclusion = sp_exclusion[~global_mask]

            profile_parties = set(sp_profile['Normalized Party'].unique()) if not sp_profile.empty else set()
            filter_parties = set()
            if not all_change_format_filters.empty and 'SALES_PERSON' in all_change_format_filters.columns:
                sp_filters = all_change_format_filters[all_change_format_filters['SALES_PERSON'] == salesperson]
                if 'NORMALIZED PARTY' in sp_filters.columns:
                    filter_parties = set(sp_filters['NORMALIZED PARTY'].unique())

            all_target_parties = sorted(list(profile_parties | filter_parties))

            for norm_party in all_target_parties:
                if TEST_MODE_PARTY and norm_party not in normalized_test_parties:
                    continue

                if not sp_profile.empty and norm_party in profile_parties:
                    party_rows = sp_profile[sp_profile['Normalized Party'] == norm_party]
                else:
                    party_rows = pd.DataFrame({
                        'Salesperson': [salesperson],
                        'Normalized Party': [norm_party],
                        'SoldParty': [norm_party],
                        'Country': [''],
                        'City': [''],
                        'Super Tag': ['DUMMY_TAG'],
                        'Count': [0],
                        'Match_Logic_Type': ['STANDARD'],
                        'Types_Found': [''],
                        'BidSaleCount': [0]
                    })

                result, party_display_name, country, city = process_party_from_profile(
                    party_rows, stocks_df, top_n, all_change_format_filters,
                    tolerance_level, tolerance_df, salesperson, stocks_by_tag, stocks_by_key, stocks_by_shape
                )

                if not result.empty:
                    if ENABLE_CLEAN_FILE:
                        if not global_exclusion_rules.empty:
                            result = apply_exclusion_rules(result, global_exclusion_rules)
                        if not party_specific_exclusion.empty:
                            specific_rules = party_specific_exclusion[party_specific_exclusion['NORMALIZED PARTY'] == norm_party]
                            if not specific_rules.empty:
                                result = apply_exclusion_rules(result, specific_rules)

                    if not result.empty:
                        if norm_party not in salesperson_party_reports:
                            salesperson_party_reports[norm_party] = {
                                'dataframes': [],
                                'display_name': party_display_name,
                                'country': country,
                                'city': city
                            }
                        salesperson_party_reports[norm_party]['dataframes'].append(result)

            if not salesperson_party_reports:
                 continue

            if GENERATE_MERGED_MASTER:
                for norm_party, party_info in salesperson_party_reports.items():
                    all_results_for_party = pd.concat(party_info['dataframes'], ignore_index=True)
                    if not all_results_for_party.empty:
                        master_chunk = all_results_for_party.copy()
                        master_chunk.insert(0, 'Salesperson', salesperson)
                        master_chunk.insert(1, 'Sold Party', party_info['display_name'])
                        all_consolidated_data.append(master_chunk)

            if GENERATE_INDIVIDUAL_REPORTS:
                with BytesIO() as buffer:
                    wb = Workbook()
                    wb.remove(wb.active)
                    sheets_for_nav = []

                    # FAST NATIVE PANDAS MEMORY FOR SEPARATED ZIPS
                    separated_zips_data = {}

                    for norm_party, party_info in salesperson_party_reports.items():
                        all_results_for_party = pd.concat(party_info['dataframes'], ignore_index=True)
                        all_results_for_party.drop_duplicates(subset=['Serial No'], keep='first', inplace=True)

                        # FAST ADVANCED SORTING
                        if ENABLE_ADVANCED_SORTING and sales_info_for_sorting is not None:
                            if len(all_results_for_party) >= SORT_COUNT_THRESHOLD:
                                all_results_for_party = apply_advanced_sorting_on_dataframe(
                                    all_results_for_party,
                                    party_info['display_name'],
                                    sales_info_for_sorting
                                )

                        # FAST SMALL SHEET DELETE
                        if ENABLE_SMALL_OUTPUT_DELETE and len(all_results_for_party) <= SMALL_OUTPUT_THRESHOLD:
                            continue

                        if not all_results_for_party.empty:
                            if 'Sr.No' not in all_results_for_party.columns:
                                all_results_for_party.insert(0, 'Sr.No', range(1, 1 + len(all_results_for_party)))

                            valid_sheet_name = re.sub(r'[\\/*?:"<>|]', '_', party_info['display_name'])[:31]
                            if not valid_sheet_name.strip():
                                valid_sheet_name = "PartySheet"

                            original_name = valid_sheet_name
                            count = 1
                            while valid_sheet_name in [s[0] for s in sheets_for_nav]:
                                valid_sheet_name = f"{original_name[:28]}_{count}"
                                count += 1

                            ws = wb.create_sheet(title=valid_sheet_name)
                            df_output = all_results_for_party.copy()

                            if isinstance(salesperson, str) and "aakash" in salesperson.lower() and "parmar" in salesperson.lower():
                                content_cols = [c for c in DESIRED_COLUMN_SEQUENCE_AAKASH if c in df_output.columns]
                                helper_cols = ["Is_Nearby", "Is_Filter_Match", "Match_Type", "Tolerance_Type", "Is_Pair_Protected", "DET_LINK"]
                                existing_helpers = [c for c in helper_cols if c in df_output.columns]
                                df_output = df_output[content_cols + existing_helpers]

                            headers = df_output.columns.tolist()
                            ws.append(headers)

                            for row in df_output.itertuples(index=False):
                                ws.append(list(row))

                            # Apply native links to the Main Worksheet
                            format_native_hyperlinks(ws, df_output)

                            if ENABLE_STYLING:
                                auto_adjust_column_widths(ws, df_data=df_output)
                                apply_styles_to_worksheet(ws)

                            for col_name in ["Is_Nearby", "Is_Filter_Match", "Match_Type", "Tolerance_Type", "Is_Pair_Protected", "DET_LINK", "Extracted_Pair_ID", "Pair_Comment_Found", "Pair_Row_Added", "Pair_Source"]:
                                headers_ws = [cell.value for cell in ws[1]]
                                if col_name in headers_ws:
                                    ws.delete_cols(headers_ws.index(col_name) + 1)

                            ws.freeze_panes = 'A2'
                            ws.cell(row=1, column=1).hyperlink = "#'Navigation'!A1"
                            is_top_customer = ENABLE_TOP_CUSTOMERS_HIGHLIGHT and (norm_party in top_customers_parties)
                            sheets_for_nav.append(
                                (valid_sheet_name, party_info['display_name'], ws.max_row - 1,
                                party_info['country'], party_info['city'], is_top_customer)
                            )

                            # FAST SEPARATED FILE GENERATION (NO COPYING OVERHEAD)
                            if ENABLE_OUTPUT_FILE_SEPARATED:
                                sep_wb = Workbook()
                                sep_ws = sep_wb.active
                                sep_ws.title = valid_sheet_name # Keeps the 31-char limit for the Excel Tab
                                sep_ws.append(headers)

                                for row in df_output.itertuples(index=False):
                                    sep_ws.append(list(row))

                                # Apply native links to the Separated Worksheet
                                format_native_hyperlinks(sep_ws, df_output)

                                if ENABLE_STYLING:
                                    auto_adjust_column_widths(sep_ws, df_data=df_output)
                                    apply_styles_to_worksheet(sep_ws)

                                for col_name in ["Is_Nearby", "Is_Filter_Match", "Match_Type", "Tolerance_Type", "Is_Pair_Protected", "DET_LINK", "Extracted_Pair_ID", "Pair_Comment_Found", "Pair_Row_Added", "Pair_Source"]:
                                    headers_sep = [cell.value for cell in sep_ws[1]]
                                    if col_name in headers_sep:
                                        sep_ws.delete_cols(headers_sep.index(col_name) + 1)

                                sep_ws.freeze_panes = 'A2'

                                # Create a full-length safe file name for the actual file
                                safe_full_file_name = re.sub(r'[\\/*?:"<>|]', '_', party_info['display_name']).strip()

                                with BytesIO() as sep_buffer:
                                    sep_wb.save(sep_buffer)
                                    # Saves using the full name
                                    separated_zips_data[f"{safe_full_file_name}.xlsx"] = sep_buffer.getvalue()

                    # WRITE SEPARATED FILES TO ZIP
                    if ENABLE_OUTPUT_FILE_SEPARATED and separated_zips_data:
                        clean_sp_folder = "".join(x for x in str(salesperson) if x.isalnum() or x in (" ", "_")).strip()[:50] or "Salesperson"
                        with BytesIO() as sub_zip_buffer:
                            with zipfile.ZipFile(sub_zip_buffer, 'w', zipfile.ZIP_DEFLATED) as sub_zip:
                                for fname, bdata in separated_zips_data.items():
                                    sub_zip.writestr(fname, bdata)
                            zipf.writestr(f"{clean_sp_folder}_Separated_Parties.zip", sub_zip_buffer.getvalue())

                    if sheets_for_nav:
                        sheets_for_nav.sort(key=lambda x: (str(x[3]), str(x[4])))
                        nav_ws = wb.create_sheet(title="Navigation", index=0)
                        nav_headers = ["Go To Sheet", "Party Name", "Count", "Country", "City", "Go →"]

                        nav_ws.append(nav_headers)

                        for r, item in enumerate(sheets_for_nav, 2):
                            sheet_name, display_name, count_rows, country, city, is_top_customer = item
                            nav_ws.cell(row=r, column=1, value=sheet_name)
                            nav_ws.cell(row=r, column=2, value=display_name)
                            nav_ws.cell(row=r, column=3, value=count_rows)
                            nav_ws.cell(row=r, column=4, value=country)
                            nav_ws.cell(row=r, column=5, value=city)
                            cell = nav_ws.cell(row=r, column=6, value="Go →")
                            cell.hyperlink = f"#'{sheet_name}'!A1"
                            if is_top_customer:
                                for c_idx in range(1, 7):
                                    nav_ws.cell(row=r, column=c_idx).fill = top_customers_fill

                        if ENABLE_STYLING:
                            auto_adjust_column_widths(nav_ws)
                            apply_styles_to_worksheet(nav_ws)
                        nav_ws.freeze_panes = 'A2'

                    wb.save(buffer)
                    clean_sp = "".join(x for x in str(salesperson) if x.isalnum() or x in (" ", "_")).strip()[:50] or "Report"
                    fname = f"{clean_sp}.xlsx"
                    zipf.writestr(fname, buffer.getvalue())
                    print(f"Saved: {fname}")

        if GENERATE_MERGED_MASTER and all_consolidated_data:
            print("Generating Merged Master File...")
            master_df = pd.concat(all_consolidated_data, ignore_index=True)
            all_consolidated_data = []

            with BytesIO() as master_buffer:
                master_wb = Workbook()
                master_wb.remove(master_wb.active)

                num_chunks = (len(master_df) // MAX_ROWS_PER_SHEET) + 1
                for i in range(num_chunks):
                    start_row = i * MAX_ROWS_PER_SHEET
                    end_row = (i + 1) * MAX_ROWS_PER_SHEET
                    chunk = master_df.iloc[start_row:end_row].copy()
                    if chunk.empty: continue

                    ws_name = f"Master_Data_Part_{i+1}"
                    ws = master_wb.create_sheet(title=ws_name)

                    cols_to_exclude = ['Is_Pair_Protected', 'Safe_Serial', 'Extracted_Pair_ID', 'DET_LINK']
                    cols_to_write = [c for c in chunk.columns if c not in cols_to_exclude]

                    ws.append(cols_to_write)

                    chunk_clean = chunk[cols_to_write]
                    for row in chunk_clean.itertuples(index=False):
                        ws.append(list(row))

                    # Apply native links to the Master File Worksheet
                    format_native_hyperlinks(ws, chunk_clean)

                    if ENABLE_STYLING:
                        auto_adjust_column_widths(ws, df_data=chunk_clean)
                        apply_styles_to_worksheet(ws)

                    del chunk
                    del chunk_clean

                master_wb.save(master_buffer)
                zipf.writestr("MASTER_CONSOLIDATED_REPORT.xlsx", master_buffer.getvalue())

    print(f"All reports saved to {output_zip_path}")

if __name__ == "__main__":
    base_project_path = "/content/drive/MyDrive/Diamond_Project/Algo 2.0"
    base_project_input_path = "/content/drive/MyDrive/Diamond_Project/Algo 2.0/HK Input data"

    #base_project_path = "/content/drive/MyDrive/Temp Project Test"
    #base_project_input_path = "/content/drive/MyDrive/Temp Project Test"

    #base_project_path = "/content/drive/MyDrive/Unassign_Regular_Stock"
    #base_project_input_path = "/content/drive/MyDrive/Unassign_Regular_Stock"

    #base_project_path = "/content/drive/MyDrive/Diamond_Project/Regular_Suggested"
    #base_project_input_path = "/content/drive/MyDrive/Diamond_Project/Regular_Suggested"

    #base_project_path = "/content/drive/MyDrive/test"
    #base_project_input_path = "/content/drive/MyDrive/test"

    #base_project_path = "/content/drive/MyDrive/Diamond_Project/Silver_Gold_Platinum parties output"
    #base_project_input_path = "/content/drive/MyDrive/Diamond_Project/Silver_Gold_Platinum parties output"

    stock_file = f"{base_project_path}/stock Actual.xlsx"
    filter_file = f"{base_project_path}/filters.xlsx" #filters - EBID" REGULAR
    clean_file = f"{base_project_path}/no_need.xlsx"
    sales_history_path = f"{base_project_path}/Sales_And_Actions_Folder"

    party_profile_perm_path = f"{base_project_input_path}/PARTY_PROFILE_PERM.parquet"
    party_profile_temp_path = f"{base_project_input_path}/PARTY_PROFILE_TEMP.parquet"
    output_zip = f"{base_project_path}/Suggest_BID265_Stock_Master_{datetime.now().strftime('%Y-%m-%d')}.zip"

    generate_salesperson_reports_from_profile(
        stock_file, party_profile_perm_path, party_profile_temp_path,
        output_zip, filter_file, clean_file, sales_history_path, TOP_N_SUPERTAGS, TOLERANCE_LEVEL
    )
