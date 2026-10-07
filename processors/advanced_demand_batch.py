from __future__ import annotations

import hashlib
import os
import re
import sqlite3
import tempfile
import zipfile
from collections import defaultdict

import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook

FILTER_RULE_BATCH_SIZE = 5000


def _mangle_duplicate_headers(headers):
    counts = defaultdict(int)
    result = []
    for raw in headers:
        base = "" if raw is None else str(raw).strip().upper()
        count = counts[base]
        result.append(base if count == 0 else f"{base}.{count}")
        counts[base] += 1
    return result


def _clean_excel_cell(value):
    return "" if value is None else value


def _stable_rule_digest(rule, ordered_columns):
    digest = hashlib.blake2b(digest_size=16)
    for col in ordered_columns:
        value = rule.get(col, "")
        if value is None or (isinstance(value, float) and pd.isna(value)):
            text = ""
        elif isinstance(value, float) and value.is_integer():
            text = str(int(value))
        else:
            text = str(value)
        payload = text.encode("utf-8", errors="surrogatepass")
        digest.update(len(payload).to_bytes(4, "big", signed=False))
        digest.update(payload)
    return digest.digest()


class DemandCandidateIndex:
    """Lazy stock-side indexes for narrowing INCLUDE rules without changing rule semantics."""

    def __init__(self, module, stocks_df):
        self.m = module
        self.df = stocks_df
        self.text_indexes = {}
        self.numeric_indexes = {}
        self.numeric_series = {}
        self.fancy_indexes = {}
        self.measurement_series = None

    def _build_text_index(self, stock_col, apply_flo_map=False):
        key = (stock_col, bool(apply_flo_map))
        if key in self.text_indexes:
            return self.text_indexes[key]
        series = self.df[stock_col].fillna("").astype(str).str.strip().str.upper()
        if apply_flo_map:
            series = series.replace(["VSL", "SL"], "FNT")
        postings = {}
        for value, labels in series.groupby(series, sort=False).groups.items():
            postings[str(value)] = np.asarray(list(labels), dtype=np.int64)
        self.text_indexes[key] = postings
        return postings

    def text_candidates(self, stock_col, values, apply_flo_map=False):
        postings = self._build_text_index(stock_col, apply_flo_map=apply_flo_map)
        arrays = [postings.get(str(value).strip().upper()) for value in values]
        arrays = [arr for arr in arrays if arr is not None and len(arr)]
        if not arrays:
            return np.empty(0, dtype=np.int64)
        if len(arrays) == 1:
            return arrays[0]
        return np.unique(np.concatenate(arrays))

    def fancy_candidates(self, stock_col, values):
        if stock_col not in self.fancy_indexes:
            series = self.df[stock_col].map(self.m.normalize_fancy_color_value)
            postings = {}
            for value, labels in series.groupby(series, sort=False).groups.items():
                postings[str(value)] = np.asarray(list(labels), dtype=np.int64)
            self.fancy_indexes[stock_col] = postings
        postings = self.fancy_indexes[stock_col]
        arrays = [postings.get(str(value).strip().upper()) for value in values]
        arrays = [arr for arr in arrays if arr is not None and len(arr)]
        if not arrays:
            return np.empty(0, dtype=np.int64)
        if len(arrays) == 1:
            return arrays[0]
        return np.unique(np.concatenate(arrays))

    def _build_numeric_index(self, stock_col):
        if stock_col in self.numeric_indexes:
            return self.numeric_indexes[stock_col]
        series = pd.to_numeric(self.df[stock_col], errors="coerce")
        self.numeric_series[stock_col] = series
        values = series.to_numpy(dtype=float, copy=False)
        labels = self.df.index.to_numpy(dtype=np.int64, copy=False)
        valid = ~np.isnan(values)
        valid_values = values[valid]
        valid_labels = labels[valid]
        order = np.argsort(valid_values, kind="mergesort")
        payload = (valid_values[order], valid_labels[order], labels[~valid])
        self.numeric_indexes[stock_col] = payload
        return payload

    @staticmethod
    def _numeric_part_candidates(sorted_values, sorted_labels, nan_labels, part):
        text = str(part).strip().upper()
        try:
            if "TO" in text:
                pieces = text.split("TO")
                if len(pieces) != 2:
                    return np.empty(0, dtype=np.int64)
                low = float(pieces[0].strip())
                high = float(pieces[1].strip())
                left = np.searchsorted(sorted_values, low, side="left")
                right = np.searchsorted(sorted_values, high, side="right")
                return sorted_labels[left:right]
            if text.startswith(">="):
                value = float(text[2:].strip())
                left = np.searchsorted(sorted_values, value, side="left")
                return sorted_labels[left:]
            if text.startswith(">"):
                value = float(text[1:].strip())
                left = np.searchsorted(sorted_values, value, side="right")
                return sorted_labels[left:]
            if text.startswith("<="):
                value = float(text[2:].strip())
                right = np.searchsorted(sorted_values, value, side="right")
                return sorted_labels[:right]
            if text.startswith("<"):
                value = float(text[1:].strip())
                right = np.searchsorted(sorted_values, value, side="left")
                return sorted_labels[:right]
            if text.startswith("!="):
                value = float(text[2:].strip())
                left = np.searchsorted(sorted_values, value, side="left")
                right = np.searchsorted(sorted_values, value, side="right")
                arrays = [arr for arr in (sorted_labels[:left], sorted_labels[right:], nan_labels) if len(arr)]
                return np.unique(np.concatenate(arrays)) if arrays else np.empty(0, dtype=np.int64)
            value = float(text)
            left = np.searchsorted(sorted_values, value, side="left")
            right = np.searchsorted(sorted_values, value, side="right")
            return sorted_labels[left:right]
        except Exception:
            return np.empty(0, dtype=np.int64)

    def numeric_candidates(self, stock_col, filter_string, allow_no_dot=False):
        sorted_values, sorted_labels, nan_labels = self._build_numeric_index(stock_col)
        text = str(filter_string).strip().upper()
        no_dot = bool(allow_no_dot and "NO DOT" in text)
        if no_dot:
            text = text.replace("NO DOT", "").strip()

        if not text:
            candidates = self.df.index.to_numpy(dtype=np.int64, copy=False)
        else:
            arrays = []
            for part in [item.strip() for item in text.split(",") if item.strip()]:
                arr = self._numeric_part_candidates(sorted_values, sorted_labels, nan_labels, part)
                if len(arr):
                    arrays.append(arr)
            candidates = np.unique(np.concatenate(arrays)) if arrays else np.empty(0, dtype=np.int64)

        if no_dot and len(candidates):
            series = self.numeric_series.get(stock_col)
            if series is None:
                series = pd.to_numeric(self.df[stock_col], errors="coerce")
                self.numeric_series[stock_col] = series
            allowed = ~series.astype(str).str.endswith(".0", na=False)
            allowed_labels = set(series.index[allowed].tolist())
            candidates = np.asarray([idx for idx in candidates if int(idx) in allowed_labels], dtype=np.int64)
        return candidates

    def measurement_candidates(self, stock_col, filter_string):
        if self.measurement_series is None:
            self.measurement_series = pd.to_numeric(
                self.df[stock_col].astype(str).str.rsplit("*", n=1).str[-1].str.strip(),
                errors="coerce",
            )
        values = self.measurement_series.to_numpy(dtype=float, copy=False)
        labels = self.df.index.to_numpy(dtype=np.int64, copy=False)
        valid = ~np.isnan(values)
        valid_values = values[valid]
        valid_labels = labels[valid]
        order = np.argsort(valid_values, kind="mergesort")
        sorted_values = valid_values[order]
        sorted_labels = valid_labels[order]
        nan_labels = labels[~valid]
        arrays = []
        for part in [item.strip() for item in str(filter_string).split(",") if item.strip()]:
            arr = self._numeric_part_candidates(sorted_values, sorted_labels, nan_labels, part)
            if len(arr):
                arrays.append(arr)
        return np.unique(np.concatenate(arrays)) if arrays else np.empty(0, dtype=np.int64)

    def candidates_for_rule(self, rule, tolerance_level):
        m = self.m
        if str(rule.get("REMARKS", "INCLUDE")).strip().upper() == "EXCLUDE":
            return None

        column_mapping = {
            "SHP": "Shp", "QUA": "Qua", "CUT": "Cut", "POL": "Pol", "SYM": "Sym",
            "FLO": "Flo", "LAB": "Lab", "DISC%": "Disc%", "SHADE": "Shade",
            "MILKY": "Milky", "E.FACET": "E.Facet", "TOIN_NAME": "TOIN_NAME",
            "SOIN_NAME": "SOIN_NAME", "GIRDLE_OPEN_NAME": "GIRDLE_OPEN_NAME",
            "PAV_OPEN_NAME": "PAV_OPEN_NAME", "TBINC_NAME": "TBINC_NAME",
            "SBINC_NAME": "SBINC_NAME", "MEASUREMENT": "Measurement",
            "REMARKS": "Remarks", "COLOR NAME": "Color", "CARAT": "Carat",
            "RESERVE%": "Disc%", "RATIO": "Ratio", "RATE": "Rate", "RAPRATE": "RapRate",
            "AMOUNT": "Amount", "TABLE INCL": "Table Incl", "DEP": "Dep", "TAB": "Tab",
            "C/A": "C/A", "C/H": "C/H", "P/A": "P/A", "P/H": "P/H", "GIRDLE": "Girdle",
            "DIA": "Dia", "SGP COMMENT": "SGPComment", "SGPCOMMENT": "SGPComment",
            "KEYTOSYMBOL": "KeyToSymbol", "COMMENT": "comment", "COLOR DESC": "Color Desc",
            "FANCY COLOR": "Fancy Color", "FANCYCOLOR": "Fancy Color",
            "EYE CLEAN": "EYE Clean", "LUSTER": "Luster", "CULET": "Culet",
        }

        best = None
        saw_condition = False
        for filter_col, filter_value in rule.items():
            filter_key = m.normalize_column_key(filter_col)
            if filter_key in m.METADATA_FILTER_KEYS or pd.isna(filter_value) or str(filter_value).strip() == "":
                continue

            if filter_key == "FANCYCOLOR":
                stock_col = m.build_column_lookup(self.df.columns).get("FANCYCOLOR")
                if not stock_col and "Color" in self.df.columns:
                    stock_col = "Color"
            else:
                stock_col = m.resolve_stock_column(filter_col, self.df, column_mapping)
            if not stock_col:
                continue

            saw_condition = True
            val_str = str(filter_value).strip().upper()
            candidate = None

            if filter_col in ["SHP", "LAB"]:
                acc = m.parse_filter_value(val_str, tolerance_level=0)
                candidate = self.text_candidates(stock_col, acc) if acc else np.empty(0, dtype=np.int64)
            elif filter_col == "CARAT":
                candidate = self.numeric_candidates(stock_col, val_str, allow_no_dot=True)
            elif filter_col == "MEASUREMENT":
                candidate = self.measurement_candidates(stock_col, val_str)
            elif filter_col in m.ORDER_MAPPING:
                order = m.ORDER_MAPPING[filter_col]
                apply_map = bool(m.ENABLE_FLO_MAPPING and filter_col == "FLO")
                effective_tolerance = (
                    tolerance_level
                    if filter_key != "FANCYCOLOR" or m.ENABLE_FANCY_COLOR_TOLERANCE
                    else 0
                )
                acc = m.parse_filter_value(
                    val_str, order, effective_tolerance, apply_flo_mapping=apply_map
                )
                if filter_key == "FANCYCOLOR":
                    candidate = self.fancy_candidates(stock_col, acc) if acc else np.empty(0, dtype=np.int64)
                else:
                    candidate = self.text_candidates(stock_col, acc, apply_flo_map=apply_map) if acc else np.empty(0, dtype=np.int64)
            elif filter_key in m.LONG_TEXT_COLUMN_KEYS:
                candidate = None
            elif filter_key in m.NUMERIC_FILTER_KEYS or m.is_numeric_stock_column(self.df[stock_col]):
                candidate = self.numeric_candidates(stock_col, val_str, allow_no_dot=False)
            else:
                acc = m.parse_filter_value(val_str, tolerance_level=0)
                candidate = self.text_candidates(stock_col, acc) if acc else np.empty(0, dtype=np.int64)

            if candidate is not None:
                if len(candidate) == 0:
                    return candidate
                if best is None or len(candidate) < len(best):
                    best = candidate

        if not saw_condition:
            return np.empty(0, dtype=np.int64)
        return best


def _order_output_columns(module, df):
    if df.empty:
        return df
    current_cols = df.columns.tolist()
    ordered_cols = []
    active_sequence = module.DESIRED_COLUMN_SEQUENCE.copy()
    if module.ENABLE_RELATED_TAG and "Tag_No" not in active_sequence:
        active_sequence.append("Tag_No")
    for col in active_sequence:
        if col in current_cols and col not in ordered_cols:
            ordered_cols.append(col)
    for col in current_cols:
        if col not in ordered_cols:
            ordered_cols.append(col)
    return df[ordered_cols]


def generate_salesperson_reports_batched(
    module,
    file_path_stock,
    output_zip_path,
    filter_file_path,
    tolerance_level,
    clean_file_path=None,
    rule_batch_size=FILTER_RULE_BATCH_SIZE,
):
    if not filter_file_path or not os.path.exists(filter_file_path):
        raise FileNotFoundError("Filters workbook is required.")
    if not file_path_stock or not os.path.exists(file_path_stock):
        raise FileNotFoundError("Stock Actual workbook is required.")

    print("Loading Stock Actual...")
    stocks_df = pd.read_excel(file_path_stock, sheet_name="Sheet1", keep_default_na=False)
    if "Serial No" in stocks_df.columns:
        stocks_df.drop_duplicates(subset=["Serial No"], keep="first", inplace=True)

    reverse_stock_map = {}
    for final_name, variations in module.STOCK_COLUMN_MAPPING.items():
        for variation in variations:
            reverse_stock_map[variation.upper()] = final_name
    stocks_df.columns = [reverse_stock_map.get(col.strip().upper(), col.strip()) for col in stocks_df.columns]
    if "Serial No" not in stocks_df.columns:
        raise ValueError("Stock workbook must contain a 'Serial No' column.")

    with tempfile.TemporaryDirectory(prefix="advanced_demand_batch_") as temp_dir:
        db_path = os.path.join(temp_dir, "matches.sqlite3")
        conn = sqlite3.connect(db_path)
        conn.execute("PRAGMA journal_mode=OFF")
        conn.execute("PRAGMA synchronous=OFF")
        conn.execute("PRAGMA temp_store=FILE")
        conn.execute(
            "CREATE TABLE matches ("
            " salesperson TEXT NOT NULL, norm_party TEXT NOT NULL, stock_idx INTEGER NOT NULL,"
            " match_rank INTEGER NOT NULL, tag_no,"
            " PRIMARY KEY (salesperson, norm_party, stock_idx)) WITHOUT ROWID"
        )
        conn.execute(
            "CREATE TABLE party_meta ("
            " salesperson TEXT NOT NULL, norm_party TEXT NOT NULL, display_name TEXT,"
            " country TEXT, city TEXT, first_seq INTEGER NOT NULL,"
            " PRIMARY KEY (salesperson, norm_party)) WITHOUT ROWID"
        )
        conn.execute(
            "CREATE TABLE salesperson_meta (salesperson TEXT PRIMARY KEY, first_seq INTEGER NOT NULL) WITHOUT ROWID"
        )
        conn.execute(
            "CREATE TABLE seen_rules (digest BLOB PRIMARY KEY) WITHOUT ROWID"
        )
        conn.execute("CREATE INDEX idx_party_meta_sp ON party_meta(salesperson, first_seq)")

        stock_cache = module.build_stock_filter_cache(stocks_df)
        candidate_index = DemandCandidateIndex(module, stocks_df)

        filter_wb = load_workbook(filter_file_path, read_only=True, data_only=True)
        try:
            removed_parties = set()
            if "Removed Party" in filter_wb.sheetnames:
                ws_removed = filter_wb["Removed Party"]
                removed_rows_iter = ws_removed.iter_rows(values_only=True)
                header = next(removed_rows_iter, ())
                header_lookup = {
                    str(value).strip().upper(): index
                    for index, value in enumerate(header)
                    if value is not None
                }
                sold_idx = header_lookup.get("SOLD PARTY")
                if sold_idx is not None:
                    for row in removed_rows_iter:
                        if sold_idx < len(row) and row[sold_idx] not in (None, ""):
                            removed_parties.add(module.normalize_party_name(str(row[sold_idx])))
            print(f"Removed Party rules loaded: {len(removed_parties):,}")

            if "Change Format" not in filter_wb.sheetnames:
                raise ValueError("Filters workbook must contain a 'Change Format' sheet.")

            ws = filter_wb["Change Format"]
            row_iter = ws.iter_rows(values_only=True)
            raw_header = next(row_iter, None)
            if raw_header is None:
                raise ValueError("Filters workbook 'Change Format' sheet is empty.")
            headers = _mangle_duplicate_headers(raw_header)
            if "SOLD PARTY" not in headers or "SALES_PERSON" not in headers:
                raise ValueError("'Change Format' must contain 'Sold Party' and 'Sales_Person' columns.")

            dedup_columns = [
                col for col in headers
                if module.ENABLE_RELATED_TAG or col not in {"TAG_NO", "TAG"}
            ] + ["NORMALIZED PARTY"]
            removed_filter_rows = 0
            duplicate_filter_rows = 0
            blank_filter_rows = 0
            processed_filter_rows = 0
            matched_filter_rows = 0
            matched_stock_links = 0
            first_match_seq = 0
            pending_since_commit = 0
            rank_map = {"Exact": 0, "Nearby": 1, "Maybe": 2}

            print(f"Streaming Change Format in batches of {int(rule_batch_size):,} rules...")
            for raw_row in row_iter:
                values = [_clean_excel_cell(value) for value in raw_row]
                if len(values) < len(headers):
                    values.extend([""] * (len(headers) - len(values)))
                elif len(values) > len(headers):
                    values = values[:len(headers)]
                if all(value == "" for value in values):
                    blank_filter_rows += 1
                    continue

                rule = dict(zip(headers, values))
                party_display_name = str(rule.get("SOLD PARTY", "")).strip()
                salesperson = str(rule.get("SALES_PERSON", "")).strip().title()
                norm_party = module.normalize_party_name(party_display_name)
                rule["SALES_PERSON"] = salesperson
                rule["NORMALIZED PARTY"] = norm_party

                if norm_party in removed_parties:
                    removed_filter_rows += 1
                    continue

                digest = _stable_rule_digest(rule, dedup_columns)
                inserted = conn.execute(
                    "INSERT OR IGNORE INTO seen_rules(digest) VALUES (?)",
                    (sqlite3.Binary(digest),),
                )
                if inserted.rowcount == 0:
                    duplicate_filter_rows += 1
                    continue
                processed_filter_rows += 1
                pending_since_commit += 1

                candidate_labels = candidate_index.candidates_for_rule(rule, tolerance_level)
                if candidate_labels is not None:
                    if len(candidate_labels) == 0:
                        if pending_since_commit >= rule_batch_size:
                            conn.commit()
                            print(f"  Processed {processed_filter_rows:,} unique rules...")
                            pending_since_commit = 0
                        continue
                    working_df = stocks_df.loc[candidate_labels]
                    working_cache = None
                else:
                    working_df = stocks_df
                    working_cache = stock_cache

                rule_series = pd.Series(rule)
                matched = module.apply_tolerance_filter(
                    working_df, rule_series, tolerance_level, stock_cache=working_cache
                )
                if matched.empty:
                    if pending_since_commit >= rule_batch_size:
                        conn.commit()
                        print(f"  Processed {processed_filter_rows:,} unique rules...")
                        pending_since_commit = 0
                    continue

                match_types = module.classify_match_types_vectorized(matched, rule_series)
                tag_value = ""
                if module.ENABLE_RELATED_TAG:
                    raw_tag = rule.get("TAG_NO", rule.get("TAG", ""))
                    tag_value = "" if pd.isna(raw_tag) else raw_tag

                first_match_seq += 1
                conn.execute(
                    "INSERT OR IGNORE INTO salesperson_meta(salesperson, first_seq) VALUES (?, ?)",
                    (salesperson, first_match_seq),
                )
                conn.execute(
                    "INSERT OR IGNORE INTO party_meta("
                    " salesperson, norm_party, display_name, country, city, first_seq)"
                    " VALUES (?, ?, ?, ?, ?, ?)",
                    (
                        salesperson,
                        norm_party,
                        party_display_name,
                        str(rule.get("COUNTRY", "")),
                        str(rule.get("CITY", "")),
                        first_match_seq,
                    ),
                )

                match_rows = [
                    (
                        salesperson,
                        norm_party,
                        int(stock_idx),
                        rank_map.get(str(match_type), 2),
                        tag_value,
                    )
                    for stock_idx, match_type in match_types.items()
                ]
                conn.executemany(
                    "INSERT INTO matches(salesperson, norm_party, stock_idx, match_rank, tag_no)"
                    " VALUES (?, ?, ?, ?, ?)"
                    " ON CONFLICT(salesperson, norm_party, stock_idx) DO UPDATE SET"
                    " match_rank = CASE WHEN excluded.match_rank < matches.match_rank"
                    " THEN excluded.match_rank ELSE matches.match_rank END,"
                    " tag_no = CASE WHEN excluded.match_rank < matches.match_rank"
                    " THEN excluded.tag_no ELSE matches.tag_no END",
                    match_rows,
                )
                matched_filter_rows += 1
                matched_stock_links += len(match_rows)

                if pending_since_commit >= rule_batch_size:
                    conn.commit()
                    print(
                        f"  Processed {processed_filter_rows:,} unique rules; "
                        f"{matched_filter_rows:,} produced matches..."
                    )
                    pending_since_commit = 0
            conn.commit()
        finally:
            filter_wb.close()

        exclusion_data = pd.DataFrame()
        if module.ENABLE_CLEAN_FILE and clean_file_path and os.path.exists(clean_file_path):
            exclusion_data = module.load_exclusion_data(clean_file_path)
            print(f"No Need rules loaded: {len(exclusion_data):,}")
        elif module.ENABLE_CLEAN_FILE:
            print("No Need workbook not supplied: clean/exclusion rules skipped.")
        else:
            print("Clean/exclusion rules disabled by configuration.")
        global_exclusion_by_sp, specific_exclusion_by_sp_party = module.build_exclusion_rule_index(exclusion_data)

        reports_created = 0
        print("Building salesperson report workbooks...")
        with zipfile.ZipFile(output_zip_path, "w", zipfile.ZIP_STORED) as zipf:
            salesperson_rows = conn.execute(
                "SELECT salesperson FROM salesperson_meta ORDER BY first_seq"
            ).fetchall()
            for (salesperson,) in salesperson_rows:
                workbook = Workbook()
                workbook.remove(workbook.active)
                sheets_for_nav = []
                global_exclusion_rules = global_exclusion_by_sp.get(salesperson, pd.DataFrame())
                party_rows = conn.execute(
                    "SELECT norm_party, display_name, country, city FROM party_meta"
                    " WHERE salesperson = ? ORDER BY first_seq",
                    (salesperson,),
                ).fetchall()
                used_sheet_names = set()

                for norm_party, party_display_name, country, city in party_rows:
                    matched_rows = conn.execute(
                        "SELECT stock_idx, match_rank, tag_no FROM matches"
                        " WHERE salesperson = ? AND norm_party = ?",
                        (salesperson, norm_party),
                    ).fetchall()
                    if not matched_rows:
                        continue

                    stock_labels = [int(row[0]) for row in matched_rows]
                    rank_by_label = {int(row[0]): int(row[1]) for row in matched_rows}
                    tag_by_label = {int(row[0]): row[2] for row in matched_rows}
                    result = stocks_df.loc[stock_labels].copy()
                    reverse_rank = {0: "Exact", 1: "Nearby", 2: "Maybe"}
                    result["Match_Type"] = [
                        reverse_rank.get(rank_by_label.get(int(index), 2), "Maybe")
                        for index in result.index
                    ]
                    result["Is_Filter_Match"] = True
                    if module.ENABLE_RELATED_TAG:
                        result["Tag_No"] = [tag_by_label.get(int(index), "") for index in result.index]

                    result["Match_Type"] = pd.Categorical(
                        result["Match_Type"], categories=["Exact", "Nearby", "Maybe"], ordered=True
                    )
                    result.sort_values(
                        by=["Match_Type", "Shp", "Carat", "Color", "Qua"], inplace=True
                    )
                    result.drop_duplicates(subset=["Serial No"], keep="first", inplace=True)
                    result.reset_index(drop=True, inplace=True)

                    if module.ENABLE_CLEAN_FILE and not result.empty:
                        specific_rules = specific_exclusion_by_sp_party.get(
                            (salesperson, norm_party), pd.DataFrame()
                        )
                        if not global_exclusion_rules.empty and not specific_rules.empty:
                            combined_no_need_rules = pd.concat(
                                [global_exclusion_rules, specific_rules], ignore_index=True
                            )
                        elif not global_exclusion_rules.empty:
                            combined_no_need_rules = global_exclusion_rules
                        else:
                            combined_no_need_rules = specific_rules
                        if not combined_no_need_rules.empty:
                            result = module.apply_exclusion_rules(result, combined_no_need_rules)
                        result.reset_index(drop=True, inplace=True)

                    result = module.apply_match_status_output_filter(result, party_display_name)
                    result.reset_index(drop=True, inplace=True)
                    if result.empty:
                        continue

                    if "Sr.No" in result.columns:
                        result.drop(columns=["Sr.No"], inplace=True)
                    result.insert(0, "Sr.No", range(1, 1 + len(result)))
                    result = _order_output_columns(module, result)

                    valid_sheet_name = re.sub(r'[\\/*?:"<>|]', "_", str(party_display_name))[:31].strip() or "PartySheet"
                    original_name = valid_sheet_name
                    counter = 1
                    while valid_sheet_name in used_sheet_names:
                        valid_sheet_name = f"{original_name[:28]}_{counter}"
                        counter += 1
                    used_sheet_names.add(valid_sheet_name)

                    ws_out = workbook.create_sheet(title=valid_sheet_name)
                    df_output = result.copy()
                    ws_out.append(df_output.columns.tolist())
                    for row in df_output.itertuples(index=False, name=None):
                        ws_out.append(row)
                    module.format_native_hyperlinks(ws_out, df_output)
                    module.apply_styles_to_worksheet(ws_out, df_output)
                    module.auto_adjust_column_widths(ws_out, dataframe=df_output)
                    for hidden_col in ["Match_Type", "Is_Filter_Match"]:
                        output_headers = [cell.value for cell in ws_out[1]]
                        if hidden_col in output_headers:
                            ws_out.delete_cols(output_headers.index(hidden_col) + 1)
                    sheets_for_nav.append(
                        (valid_sheet_name, party_display_name, len(df_output), country, city)
                    )

                if not sheets_for_nav:
                    workbook.close()
                    continue

                nav_ws = workbook.create_sheet(title="Navigation", index=0)
                nav_headers = ["Go To Sheet", "Party Name", "Count", "Country", "City", "Go \u2192"]
                for col_num, header_name in enumerate(nav_headers, start=1):
                    nav_ws.cell(row=1, column=col_num, value=header_name)
                for row_num, (sheet_name, party_name, count, country, city) in enumerate(sheets_for_nav, start=2):
                    nav_ws.cell(row=row_num, column=1, value=sheet_name)
                    nav_ws.cell(row=row_num, column=2, value=party_name)
                    nav_ws.cell(row=row_num, column=3, value=count)
                    nav_ws.cell(row=row_num, column=4, value=country)
                    nav_ws.cell(row=row_num, column=5, value=city)
                    nav_ws.cell(row=row_num, column=6, value="Go \u2192").hyperlink = f"#'{sheet_name}'!A1"
                module.apply_styles_to_worksheet(nav_ws)
                module.auto_adjust_column_widths(nav_ws)

                clean_sp = "".join(
                    char for char in str(salesperson) if char.isalnum() or char in (" ", "_")
                ).strip()[:50] or "Report"
                temp_xlsx = os.path.join(temp_dir, f"{clean_sp}.xlsx")
                workbook.save(temp_xlsx)
                workbook.close()
                zipf.write(temp_xlsx, arcname=f"{clean_sp}.xlsx")
                try:
                    os.remove(temp_xlsx)
                except OSError:
                    pass
                reports_created += 1

        conn.close()

    print("PROCESS COMPLETE")
    print(f"Output file: {output_zip_path}")
    print(f"Unique filter rules processed: {processed_filter_rows:,}")
    print(f"Removed-party filter rows skipped: {removed_filter_rows:,}")
    print(f"Duplicate filter rows skipped: {duplicate_filter_rows:,}")
    print(f"Blank filter rows skipped: {blank_filter_rows:,}")
    print(f"Rules that produced matches: {matched_filter_rows:,}")
    print(f"Matched stock links before party dedupe: {matched_stock_links:,}")
    print(f"Salesperson reports created: {reports_created:,}")
