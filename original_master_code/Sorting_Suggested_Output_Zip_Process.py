import re
import math
import shutil
import zipfile
from pathlib import Path
from collections import defaultdict
from copy import copy

import pandas as pd
from openpyxl import load_workbook


# =========================================================
# CONFIG
# =========================================================

INPUT_ZIP_OR_FOLDER = "/content/drive/MyDrive/Diamond_Project/Algo 2.0/Suggest_BID265_Stock_2026-07-15.zip"
SALES_FILE = "/content/drive/MyDrive/Diamond_Project/Algo 2.0/Sales_And_Actions_Folder/Sales and bid.xlsx"

OUTPUT_FOLDER = r"/content/drive/MyDrive/Diamond_Project/Algo 2.0/Sales_And_Actions_Folder/sorted_output_fast_keep_order"
OUTPUT_ZIP_NAME = r"sorted_output_fast_keep_order.zip"

# ---------------------------------------------------------
# NEW: MATCH TYPE FILTERING
# ---------------------------------------------------------
ONLY_MATCHED_TYPE = True

# Must exactly match strings returned by detector:
# "Exact", "Near by", "May be", "Filter"
ALLOWED_MATCH_TYPES = {"Exact", "Near by", "Filter"}

# Only process parties whose Navigation Count is above this
COUNT_THRESHOLD = 200#300

# Final limits
NORMAL_PARTY_LIMIT = 100#150
HIGH_PRIORITY_PARTY_LIMIT = 150#250

# Family caps
EXACT_CAP = 3
NEAR_CAP = 2
MAYBE_CAP = 1

# If party type is one of these, use HIGH_PRIORITY_PARTY_LIMIT
HIGH_PRIORITY_PARTY_TYPES = {
    "PLATINUM", "GOLD", "DIAMOND", "PREMIUM", "VIP"
}

# Or if party total history >= this, also treat as high priority
HIGH_PRIORITY_TOTAL_HISTORY_THRESHOLD = 200

# Output row colors
GREEN_RGB_VALUES = {"FFCCFFCC", "00CCFFCC"}
ORANGE_RGB_VALUES = {"FFFFE699", "00FFE699"}

# Filter rows from your screenshot look like Excel light yellow / cream fill.
# Add/remove RGB values here if your actual workbook uses a different shade.
FILTER_RGB_VALUES = {
    "FFFFF2CC", "00FFF2CC",   # Excel light yellow
    "FFFFFFCC", "00FFFFCC",   # pale yellow fallback
    "FFFFFF99", "00FFFF99",   # another common yellow
}

PAIR_WORDS = ("PAIR", "PAIRED", "MATCH PAIR")

TYPE_WEIGHTS = {
    "SALE": 20,
    "BID": 15,
    "OFFER": 10,
    "WISHLIST": 10,
    "CART": 10,
    "HOLD": 7,
    "INQUIRY": 2,
    "ENQUIRY": 2,
    "LIST": 2,
    "VIEW": 4,
    "SEARCH": 3,
    "DETAILS": 5,
}

DEFAULT_TYPE_WEIGHT = 1

# =========================================================
# BASIC HELPERS
# =========================================================

def norm(v):
    if v is None:
        return ""
    return str(v).strip()

def norm_upper(v):
    return norm(v).upper()

def safe_float(v, default=0.0):
    try:
        if v is None:
            return default
        if isinstance(v, float) and math.isnan(v):
            return default
        return float(v)
    except Exception:
        return default

def get_fill_rgb(cell):
    fill = cell.fill
    if fill is None:
        return ""
    fg = fill.fgColor
    if fg is None:
        return ""
    rgb = getattr(fg, "rgb", None)
    return str(rgb).upper() if rgb else ""

def detect_match_type_from_row(ws, row_num, max_check_cols=5):
    rgbs = []

    for col in range(1, min(ws.max_column, max_check_cols) + 1):
        rgb = get_fill_rgb(ws.cell(row_num, col))
        if rgb:
            rgbs.append(rgb)

    if any(x in GREEN_RGB_VALUES for x in rgbs):
        return "Exact"

    if any(x in ORANGE_RGB_VALUES for x in rgbs):
        return "Near by"

    if any(x in FILTER_RGB_VALUES for x in rgbs):
        return "Filter"

    return "May be"

def is_pair_comment(text):
    t = norm_upper(text)
    return any(word in t for word in PAIR_WORDS)

def cps_score(cut, pol, sym):
    rank = {"EX": 4, "VG": 3, "G": 2, "F": 1}
    return (
        rank.get(norm_upper(cut), 0)
        + rank.get(norm_upper(pol), 0)
        + rank.get(norm_upper(sym), 0)
    )

def event_weight(event_type):
    return TYPE_WEIGHTS.get(norm_upper(event_type), DEFAULT_TYPE_WEIGHT)

def build_family_key(shape, carat_grp, color, qua, lab):
    return "|".join([
        norm_upper(shape),
        norm(carat_grp),
        norm_upper(color),
        norm_upper(qua),
        norm_upper(lab),
    ])

def build_variant_key(cut, pol, sym, flo):
    return "|".join([
        norm_upper(cut),
        norm_upper(pol),
        norm_upper(sym),
        norm_upper(flo),
    ])

# =========================================================
# SALES HISTORY
# =========================================================

def parse_carat_ranges_from_sales(sales_df):
    ranges = []
    if "Carat_grp" not in sales_df.columns:
        return ranges

    seen = set()
    for grp in sales_df["Carat_grp"].dropna().astype(str).unique():
        grp = grp.strip()
        if grp in seen:
            continue
        seen.add(grp)

        m = re.match(r"^\s*(\d+(?:\.\d+)?)\s*-\s*(\d+(?:\.\d+)?)\s*$", grp)
        if m:
            low = float(m.group(1))
            high = float(m.group(2))
            ranges.append((low, high, grp))

    ranges.sort(key=lambda x: (x[0], x[1]))
    return ranges

def derive_carat_group(carat_value, parsed_ranges):
    try:
        carat = float(carat_value)
    except Exception:
        return ""

    for low, high, label in parsed_ranges:
        if low <= carat <= high:
            return label

    if parsed_ranges:
        nearest = min(parsed_ranges, key=lambda x: min(abs(carat - x[0]), abs(carat - x[1])))
        return nearest[2]
    return ""

def load_sales_history(sales_file):
    sales = pd.read_excel(sales_file)
    sales.columns = [str(c).strip() for c in sales.columns]

    required = ["Shp", "Carat_grp", "Col", "Qua", "Lab", "Sold Party"]
    missing = [c for c in required if c not in sales.columns]
    if missing:
        raise ValueError(f"Missing required columns in sales.xlsx: {missing}")

    if "Type" not in sales.columns:
        sales["Type"] = ""

    parsed_ranges = parse_carat_ranges_from_sales(sales)

    sales["FAMILY_KEY"] = sales.apply(
        lambda r: build_family_key(
            r.get("Shp"),
            r.get("Carat_grp"),
            r.get("Col"),
            r.get("Qua"),
            r.get("Lab"),
        ),
        axis=1
    )

    sales["TYPE_WEIGHT"] = sales["Type"].apply(event_weight)

    party_family_map = defaultdict(float)
    global_family_map = defaultdict(float)
    party_total_map = defaultdict(float)
    party_type_map = {}

    for _, r in sales.iterrows():
        sold_party = norm_upper(r.get("Sold Party"))
        family_key = r.get("FAMILY_KEY", "")
        wt = safe_float(r.get("TYPE_WEIGHT"), 1)

        party_family_map[(sold_party, family_key)] += wt
        global_family_map[family_key] += wt
        party_total_map[sold_party] += wt

        if "Party_type" in sales.columns and sold_party and sold_party not in party_type_map:
            party_type_map[sold_party] = norm_upper(r.get("Party_type"))

    return {
        "parsed_ranges": parsed_ranges,
        "party_family_map": dict(party_family_map),
        "global_family_map": dict(global_family_map),
        "party_total_map": dict(party_total_map),
        "party_type_map": dict(party_type_map),
    }

# =========================================================
# PARTY PRIORITY
# =========================================================

def get_priority_for_party(party_name, sales_info):
    party_u = norm_upper(party_name)
    party_type = sales_info["party_type_map"].get(party_u, "")
    total_history = sales_info["party_total_map"].get(party_u, 0)

    if party_type in HIGH_PRIORITY_PARTY_TYPES:
        return "HIGH"
    if total_history >= HIGH_PRIORITY_TOTAL_HISTORY_THRESHOLD:
        return "HIGH"
    return "NORMAL"

def get_party_limit(party_name, sales_info):
    if ONLY_MATCHED_TYPE:
        return NORMAL_PARTY_LIMIT

    return HIGH_PRIORITY_PARTY_LIMIT if get_priority_for_party(party_name, sales_info) == "HIGH" else NORMAL_PARTY_LIMIT


# =========================================================
# SCORING
# =========================================================

def is_clean_row(shade, milky, sgp_comment):
    shade = norm_upper(shade)
    milky = norm_upper(milky)
    sgp = norm_upper(sgp_comment)

    bad_words = ["I.A", "OPEN", "SHADE", "MILKY", "CLOUD", "BLACK", "INCLUSION"]

    shade_ok = shade in ("", "NONE", "NO", "NIL")
    milky_ok = milky in ("", "NONE", "NO", "NIL")
    sgp_bad = any(w in sgp for w in bad_words)

    return 1 if (shade_ok and milky_ok and not sgp_bad) else 0

def score_tuple(item):
    match_rank = {"Exact": 3, "Near by": 2, "May be": 1}
    return (
        item["is_pair"],
        match_rank.get(item["match_type"], 0),
        item["party_family_count"],
        item["global_family_count"],
        item["is_clean"],
        item["cps"],
        item["disc"],
    )

def unique_keep_allowed(item):
    if item["match_type"] == "Exact":
        return True
    if item["is_pair"]:
        return True
    if item["party_family_count"] >= 1:
        return True
    if item["global_family_count"] >= 5:
        return True
    return False


# =========================================================
# EXCEL HELPERS
# =========================================================

def get_header_map(ws):
    return {
        norm(ws.cell(1, c).value): c
        for c in range(1, ws.max_column + 1)
        if norm(ws.cell(1, c).value)
    }

def clone_cell(src_cell, dst_cell):
    dst_cell.value = src_cell.value
    if src_cell.has_style:
        dst_cell._style = copy(src_cell._style)
    dst_cell.number_format = src_cell.number_format
    dst_cell.font = copy(src_cell.font)
    dst_cell.fill = copy(src_cell.fill)
    dst_cell.border = copy(src_cell.border)
    dst_cell.alignment = copy(src_cell.alignment)
    dst_cell.protection = copy(src_cell.protection)

    if src_cell.hyperlink:
        dst_cell._hyperlink = copy(src_cell.hyperlink)
    else:
        dst_cell._hyperlink = None

    if src_cell.comment:
        dst_cell.comment = copy(src_cell.comment)
    else:
        dst_cell.comment = None

def snapshot_row(ws, row_num):
    return {
        "row_num": row_num,
        "height": ws.row_dimensions[row_num].height,
        "hidden": ws.row_dimensions[row_num].hidden,
        "outlineLevel": ws.row_dimensions[row_num].outlineLevel,
        "cells": [ws.cell(row_num, c) for c in range(1, ws.max_column + 1)]
    }

def write_snapshot_row(ws, target_row, snap):
    for c in range(1, ws.max_column + 1):
        src = snap["cells"][c - 1]
        dst = ws.cell(target_row, c)
        clone_cell(src, dst)

    ws.row_dimensions[target_row].height = snap["height"]
    ws.row_dimensions[target_row].hidden = snap["hidden"]
    ws.row_dimensions[target_row].outlineLevel = snap["outlineLevel"]

def clear_trailing_rows(ws, start_row):
    """
    Physically deletes the remaining rows to completely remove
    background colors, borders, and leftover formatting.
    """
    max_r = ws.max_row
    if start_row <= max_r:
        amount_to_delete = max_r - start_row + 1
        ws.delete_rows(start_row, amount_to_delete)

def reset_srno_if_exists(ws, final_count):
    header_map = get_header_map(ws)
    sr_col = None
    for key in ("Sr.No", "SR.NO", "SR NO", "SRNO"):
        if key in header_map:
            sr_col = header_map[key]
            break

    if sr_col:
        for r in range(2, 2 + final_count):
            ws.cell(r, sr_col).value = r - 1


# =========================================================
# NAVIGATION
# =========================================================

def get_target_party_names_from_navigation(nav_ws, threshold):
    header_map = get_header_map(nav_ws)

    if "Count" not in header_map:
        return set()

    # IMPORTANT: use actual worksheet name first
    if "Go To Sheet" in header_map:
        party_col = header_map["Go To Sheet"]
    elif "Party Name" in header_map:
        party_col = header_map["Party Name"]
    else:
        return set()

    count_col = header_map["Count"]
    targets = set()

    for r in range(2, nav_ws.max_row + 1):
        party = nav_ws.cell(r, party_col).value
        cnt = nav_ws.cell(r, count_col).value
        cnt = safe_float(cnt, 0)

        if party and cnt > threshold:
            targets.add(str(party).strip())

    return targets

def update_navigation_counts(nav_ws, final_counts):
    header_map = get_header_map(nav_ws)

    if "Count" not in header_map:
        return

    # IMPORTANT: match final_counts by actual worksheet name
    if "Go To Sheet" in header_map:
        party_col = header_map["Go To Sheet"]
    elif "Party Name" in header_map:
        party_col = header_map["Party Name"]
    else:
        return

    count_col = header_map["Count"]

    for r in range(2, nav_ws.max_row + 1):
        party = nav_ws.cell(r, party_col).value
        if not party:
            continue

        party_s = str(party).strip()

        for k, v in final_counts.items():
            if norm_upper(k) == norm_upper(party_s):
                nav_ws.cell(r, count_col).value = v
                break

# =========================================================
# CORE LOGIC
# =========================================================

def build_sheet_items(ws, party_name, sales_info):
    header_map = get_header_map(ws)

    required = ["Serial No", "Shp", "Carat", "Color", "Qua", "Cut", "Pol", "Sym", "Flo", "Lab", "Disc%", "SGPComment"]
    missing = [c for c in required if c not in header_map]
    if missing:
        raise ValueError(f"Sheet '{ws.title}' missing columns: {missing}")

    items = []

    for row_num in range(2, ws.max_row + 1):
        serial_no = ws.cell(row_num, header_map["Serial No"]).value
        if serial_no in (None, ""):
            continue

        shp = ws.cell(row_num, header_map["Shp"]).value
        carat = ws.cell(row_num, header_map["Carat"]).value
        color = ws.cell(row_num, header_map["Color"]).value
        qua = ws.cell(row_num, header_map["Qua"]).value
        cut = ws.cell(row_num, header_map["Cut"]).value
        pol = ws.cell(row_num, header_map["Pol"]).value
        sym = ws.cell(row_num, header_map["Sym"]).value
        flo = ws.cell(row_num, header_map["Flo"]).value
        lab = ws.cell(row_num, header_map["Lab"]).value
        disc = ws.cell(row_num, header_map["Disc%"]).value
        sgp = ws.cell(row_num, header_map["SGPComment"]).value

        shade = ws.cell(row_num, header_map["Shade"]).value if "Shade" in header_map else ""
        milky = ws.cell(row_num, header_map["Milky"]).value if "Milky" in header_map else ""

        carat_grp = derive_carat_group(carat, sales_info["parsed_ranges"])
        family_key = build_family_key(shp, carat_grp, color, qua, lab)
        variant_key = build_variant_key(cut, pol, sym, flo)

        party_family_count = sales_info["party_family_map"].get((norm_upper(party_name), family_key), 0)
        global_family_count = sales_info["global_family_map"].get(family_key, 0)

        item = {
            "row_num": row_num,
            "serial_no": serial_no,
            "match_type": detect_match_type_from_row(ws, row_num),
            "is_pair": 1 if is_pair_comment(sgp) else 0,
            "party_family_count": safe_float(party_family_count),
            "global_family_count": safe_float(global_family_count),
            "cps": cps_score(cut, pol, sym),
            "disc": safe_float(disc),
            "family_key": family_key,
            "variant_key": variant_key,
            "is_clean": is_clean_row(shade, milky, sgp),
        }
        items.append(item)

    return items

def select_keep_rows_stable(items, party_limit=None):
    """
    Keep rows using your final logic, but preserve original row order
    among kept rows for maximum speed and workbook stability.
    """

    # ----------------------------------------------------------------
    # MATCH TYPE FILTER MODE:
    # If ONLY_MATCHED_TYPE is True, keep all allowed match types only.
    # No sales history, no family caps, no party limit.
    # ----------------------------------------------------------------
    if ONLY_MATCHED_TYPE:
        kept_items = [
            it for it in items
            if it["match_type"] in ALLOWED_MATCH_TYPES
        ]

        # Preserve original Excel row order
        kept_items = sorted(kept_items, key=lambda x: x["row_num"])

        return [it["row_num"] for it in kept_items]

    # ----------------------------------------------------------------
    # ORIGINAL LOGIC: If ONLY_MATCHED_TYPE is False, do the complex
    # Family grouping, Caps, and Variant checking.
    # ----------------------------------------------------------------
    families = defaultdict(list)
    for item in items:
        families[item["family_key"]].append(item)

    keep_row_nums = set()

    for fam_items in families.values():
        fam_items_sorted = sorted(fam_items, key=score_tuple, reverse=True)

        # unique family rule
        if len(fam_items_sorted) == 1:
            it = fam_items_sorted[0]
            if unique_keep_allowed(it):
                keep_row_nums.add(it["row_num"])
            continue

        # split by match type
        by_type = {"Exact": [], "Near by": [], "May be": [], "Filter": []}
        for it in fam_items_sorted:
            by_type[it["match_type"]].append(it)

        # exact
        chosen = 0
        seen_variants = set()
        for it in by_type["Exact"]:
            if chosen >= EXACT_CAP:
                break
            if it["is_pair"] or it["variant_key"] not in seen_variants:
                keep_row_nums.add(it["row_num"])
                seen_variants.add(it["variant_key"])
                chosen += 1

        # near
        chosen = 0
        seen_variants = set()
        for it in by_type["Near by"]:
            if chosen >= NEAR_CAP:
                break
            if it["is_pair"] or it["variant_key"] not in seen_variants:
                keep_row_nums.add(it["row_num"])
                seen_variants.add(it["variant_key"])
                chosen += 1

        # maybe
        chosen = 0
        seen_variants = set()
        for it in by_type["May be"]:
            if chosen >= MAYBE_CAP:
                break
            if it["is_pair"] or it["variant_key"] not in seen_variants:
                keep_row_nums.add(it["row_num"])
                seen_variants.add(it["variant_key"])
                chosen += 1

    # global party cap
    kept_items = [it for it in items if it["row_num"] in keep_row_nums]

    if len(kept_items) > party_limit:
        top_kept = sorted(kept_items, key=score_tuple, reverse=True)[:party_limit]
        allowed = {it["row_num"] for it in top_kept}
        kept_items = [it for it in items if it["row_num"] in allowed]

    # preserve original row order
    kept_items = sorted(kept_items, key=lambda x: x["row_num"])
    return [it["row_num"] for it in kept_items]

def rewrite_sheet_keep_order(ws, keep_row_nums):
    """
    Rebuild sheet so kept rows stay in original relative order.
    This is faster and safer than full re-sorting.
    """
    keep_row_nums = sorted(keep_row_nums)
    if not keep_row_nums:
        clear_trailing_rows(ws, 2)
        return 0

    snapshots = [snapshot_row(ws, r) for r in keep_row_nums]

    target_row = 2
    for snap in snapshots:
        write_snapshot_row(ws, target_row, snap)
        target_row += 1

    clear_trailing_rows(ws, target_row)
    reset_srno_if_exists(ws, len(snapshots))
    return len(snapshots)


# =========================================================
# PROCESS WORKBOOK
# =========================================================

def process_workbook(xlsx_path, sales_info, output_path):
    print(f"\nOpening workbook: {xlsx_path}")
    wb = load_workbook(xlsx_path)

    if "Navigation" not in wb.sheetnames:
        print("  Skipped: no Navigation sheet")
        wb.save(output_path)
        return

    nav_ws = wb["Navigation"]
    target_parties = get_target_party_names_from_navigation(nav_ws, COUNT_THRESHOLD)

    # Normalize target party names and sheet names for safe matching
    target_party_norms = {norm_upper(p) for p in target_parties}

    sheet_name_map = {
        norm_upper(ws.title): ws.title
        for ws in wb.worksheets
        if ws.title != "Navigation"
    }

    missing_sheets = sorted([
        p for p in target_parties
        if norm_upper(p) not in sheet_name_map
    ])

    if missing_sheets:
        print("  WARNING: target parties not found as sheet names:")
        for p in missing_sheets:
            print(f"    - {p}")

    if not target_parties:
        print("  No parties above threshold. Workbook unchanged.")
        wb.save(output_path)
        return

    print(f"  Target parties > {COUNT_THRESHOLD}: {len(target_parties)}")
    final_counts = {}

    for ws in wb.worksheets:
        if ws.title == "Navigation":
            continue

        if norm_upper(ws.title) not in target_party_norms:
            continue

        header_map = get_header_map(ws)
        if "Serial No" not in header_map or "SGPComment" not in header_map:
            print(f"  Skipped {ws.title}: not a party output sheet")
            continue

        print(f"  Filtering sheet: {ws.title} (rows={ws.max_row})")
        items = build_sheet_items(ws, ws.title, sales_info)

        if ONLY_MATCHED_TYPE:
            keep_rows = select_keep_rows_stable(items)
        else:
            party_limit = get_party_limit(ws.title, sales_info)
            keep_rows = select_keep_rows_stable(items, party_limit)

        final_count = rewrite_sheet_keep_order(ws, keep_rows)
        final_counts[ws.title] = final_count
        print(f"    Final rows kept: {final_count}")

    update_navigation_counts(nav_ws, final_counts)
    wb.save(output_path)
    print("  Saved workbook.")


# =========================================================
# ZIP / FOLDER HANDLING
# =========================================================

def extract_input(input_path, workdir):
    input_path = Path(input_path)
    workdir = Path(workdir)

    if workdir.exists():
        shutil.rmtree(workdir)
        workdir.mkdir(parents=True, exist_ok=True)

    if input_path.is_dir():
        for p in input_path.rglob("*"):
            target = workdir / p.relative_to(input_path)
            if p.is_dir():
                target.mkdir(parents=True, exist_ok=True)
            else:
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(p, target)
        return workdir

    if input_path.suffix.lower() == ".zip":
        with zipfile.ZipFile(input_path, "r") as zf:
            zf.extractall(workdir)
        return workdir

    raise ValueError("INPUT_ZIP_OR_FOLDER must be a folder or .zip file")

def zip_folder(folder_path, zip_path):
    folder_path = Path(folder_path)
    zip_path = Path(zip_path)

    if zip_path.exists():
        zip_path.unlink()

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for file_path in folder_path.rglob("*"):
            if file_path.is_file():
                zf.write(file_path, arcname=file_path.relative_to(folder_path))


# =========================================================
# MAIN
# =========================================================

def main():
    if ONLY_MATCHED_TYPE:
        print("ONLY_MATCHED_TYPE=True, skipping sales history loading...")

        sales_info = {
            "parsed_ranges": [],
            "party_family_map": {},
            "global_family_map": {},
            "party_total_map": {},
            "party_type_map": {},
        }
    else:
        print("Loading sales history...")

        if not Path(SALES_FILE).exists():
            raise FileNotFoundError(
                f"Sales file is required when ONLY_MATCHED_TYPE=False: {SALES_FILE}"
            )

        sales_info = load_sales_history(SALES_FILE)

    workdir = Path(OUTPUT_FOLDER) / "_work"
    outdir = Path(OUTPUT_FOLDER) / "files"

    print("Extracting input...")
    extracted_dir = extract_input(INPUT_ZIP_OR_FOLDER, workdir)

    if outdir.exists():
        shutil.rmtree(outdir)
    shutil.copytree(extracted_dir, outdir)

    xlsx_files = list(outdir.rglob("*.xlsx"))
    if not xlsx_files:
        raise ValueError("No .xlsx files found in input.")

    print(f"Found {len(xlsx_files)} workbook(s).")

    for xlsx_file in xlsx_files:
        process_workbook(xlsx_file, sales_info, xlsx_file)

    print("\nCreating output zip...")
    zip_folder(outdir, OUTPUT_ZIP_NAME)
    print(f"Done: {OUTPUT_ZIP_NAME}")

    if Path(OUTPUT_FOLDER).exists():
        shutil.rmtree(OUTPUT_FOLDER)

    print(f"Cleaned folder: {OUTPUT_FOLDER}")


if __name__ == "__main__":
    main()