"""Resumable, transaction-safe Advanced Demand Filter batches.

The rule predicate, dedupe, matched-row UPSERT, party preparation and Excel
writing below are intentionally copied from advanced_demand_batch.py, so the
source master matching engine and output layout remain unchanged.
"""
from __future__ import annotations
import json
import os
import shutil
import sqlite3
import time
import zipfile
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from contextlib import nullcontext
from collections import OrderedDict
import re

from .advanced_demand_batch import (
    DemandCandidateIndex, _mangle_duplicate_headers, _clean_excel_cell,
    _stable_rule_digest, _add_navigation_sheet_streaming,
    _add_styled_party_sheet_streaming, _add_styled_party_sheet,
    _order_output_columns, _unique_party_filename,
    STREAM_REPORT_ROW_THRESHOLD, REPORT_LAYOUT_SALESPERSON, REPORT_LAYOUT_BOTH,
    REPORT_LAYOUT_SEPARATE, REPORT_LAYOUTS,
)


def _save_state(conn, state):
    conn.execute("INSERT OR REPLACE INTO incremental_state(id, payload) VALUES (1, ?)",
                 (json.dumps(state),))
    conn.commit()


class _StagedZipWriter:
    """Write completed group artifacts atomically; final ZIP is built only once."""
    def __init__(self, job_dir):
        self.folder = Path(job_dir) / "reports"
        self.folder.mkdir(exist_ok=True)
        self.members = []

    def __enter__(self):
        return self

    def __exit__(self, *_exc):
        return False

    def write(self, src, arcname):
        target = self.folder / arcname
        partial = target.with_suffix(target.suffix + ".part")
        shutil.copyfile(src, partial)
        os.replace(partial, target)
        self.members.append(arcname)


def _load_normalized_stock(job_dir, stock_path, module):
    """Read Excel once; subsequent checkpoints reuse identical pandas values/dtypes."""
    cache = Path(job_dir) / "stock_normalized.pkl"
    if cache.is_file():
        return pd.read_pickle(cache)
    print("Loading and preparing Stock Actual (first batch only)...")
    stocks_df = pd.read_excel(stock_path, sheet_name="Sheet1", keep_default_na=False)
    if "Serial No" in stocks_df.columns:
        stocks_df.drop_duplicates(subset=["Serial No"], keep="first", inplace=True)
    reverse_stock_map = {}
    for final_name, variations in module.STOCK_COLUMN_MAPPING.items():
        for variation in variations:
            reverse_stock_map[variation.upper()] = final_name
    stocks_df.columns = [reverse_stock_map.get(col.strip().upper(), col.strip()) for col in stocks_df.columns]
    if "Serial No" not in stocks_df.columns:
        raise ValueError("Stock workbook must contain a 'Serial No' column.")
    partial = cache.with_suffix(".pkl.part")
    stocks_df.to_pickle(partial)
    os.replace(partial, cache)
    return stocks_df


def _start_db(conn):
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS matches (
          salesperson TEXT NOT NULL, norm_party TEXT NOT NULL,
          stock_idx INTEGER NOT NULL, match_rank INTEGER NOT NULL, tag_no,
          PRIMARY KEY (salesperson, norm_party, stock_idx)
        ) WITHOUT ROWID;
        CREATE TABLE IF NOT EXISTS party_meta (
          salesperson TEXT NOT NULL, norm_party TEXT NOT NULL,
          display_name TEXT, country TEXT, city TEXT, first_seq INTEGER NOT NULL,
          PRIMARY KEY (salesperson, norm_party)
        ) WITHOUT ROWID;
        CREATE TABLE IF NOT EXISTS salesperson_meta (
          salesperson TEXT PRIMARY KEY, first_seq INTEGER NOT NULL
        ) WITHOUT ROWID;
        CREATE TABLE IF NOT EXISTS seen_rules (digest BLOB PRIMARY KEY) WITHOUT ROWID;
        CREATE INDEX IF NOT EXISTS idx_party_meta_sp ON party_meta(salesperson, first_seq);
        CREATE TABLE IF NOT EXISTS incremental_state (id INTEGER PRIMARY KEY, payload TEXT NOT NULL);
    """)
    state = {
        "phase": "matching", "filter_offset": 0, "report_index": 0,
        "members": [], "used_salesperson_filenames": [],
        "removed_filter_rows": 0, "duplicate_filter_rows": 0,
        "blank_filter_rows": 0, "processed_filter_rows": 0,
        "matched_filter_rows": 0, "matched_stock_links": 0,
        "first_match_seq": 0, "cache_hits": 0,
        "reports_created": 0, "below_minimum_parties": 0,
        "separated_party_files": 0,
    }
    _save_state(conn, state)
    return state


def get_state(job_dir):
    db = Path(job_dir) / "matches.sqlite3"
    if not db.is_file():
        return None
    with sqlite3.connect(db) as conn:
        row = conn.execute("SELECT payload FROM incremental_state WHERE id=1").fetchone()
        return json.loads(row[0]) if row else None


def run_checkpoint_step(module, file_path_stock, output_zip_path, filter_file_path,
                        tolerance_level, clean_file_path=None, *, job_dir,
                        batch_rows=1000, max_salespeople=1,
                        max_batch_seconds=25,
                        min_stones_per_party=0,
                        output_layout=REPORT_LAYOUT_SALESPERSON):
    if isinstance(min_stones_per_party, bool) or not isinstance(min_stones_per_party, int) or min_stones_per_party < 0:
        raise ValueError("Minimum stones per party must be a non-negative integer.")
    if output_layout not in REPORT_LAYOUTS:
        raise ValueError(f"Unknown Advanced Demand output layout: {output_layout!r}")
    if batch_rows < 1 or max_salespeople < 1:
        raise ValueError("Batch sizes must be positive")
    include_workbooks = output_layout in {REPORT_LAYOUT_SALESPERSON, REPORT_LAYOUT_BOTH}
    include_separate = output_layout in {REPORT_LAYOUT_SEPARATE, REPORT_LAYOUT_BOTH}
    job_dir = os.fspath(job_dir)
    os.makedirs(job_dir, exist_ok=True)
    db_path = os.path.join(job_dir, "matches.sqlite3")
    initial = not os.path.exists(db_path)
    with sqlite3.connect(db_path, timeout=60) as conn:
        conn.execute("PRAGMA temp_store=FILE")
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA synchronous=NORMAL")
        try:
            state = get_state(job_dir) if not initial else None
        except sqlite3.OperationalError:
            state = None
        if state is None:
            state = _start_db(conn)
        if state["phase"] == "done":
            return state
        if state["phase"] == "packaging":
            partial = output_zip_path + ".part"
            with zipfile.ZipFile(partial, "w", zipfile.ZIP_STORED) as zipf:
                for arcname in state["members"]:
                    zipf.write(os.path.join(job_dir, "reports", arcname), arcname=arcname)
            os.replace(partial, output_zip_path)
            state["phase"] = "done"
            _save_state(conn, state)
            return state
        
        if state["phase"] == "matching":
            consumed = 0
            stopped_for_checkpoint = False
            rule_batch_size = batch_rows
            batch_started = time.monotonic()
            stocks_df = _load_normalized_stock(job_dir, file_path_stock, module)

            # INCLUDE rules normally run against a small indexed stock subset.
            # Build the expensive full-stock cache only if an EXCLUDE/long-text
            # rule actually needs a whole-stock scan. This preserves its semantics.
            stock_cache = None
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
                header_iter = ws.iter_rows(min_row=1, max_row=1, values_only=True)
                raw_header = next(header_iter, None)
                row_iter = ws.iter_rows(min_row=state["filter_offset"] + 2, values_only=True)
                if raw_header is None:
                    raise ValueError("Filters workbook 'Change Format' sheet is empty.")
                headers = _mangle_duplicate_headers(raw_header)
                if "SOLD PARTY" not in headers or "SALES_PERSON" not in headers:
                    raise ValueError("'Change Format' must contain 'Sold Party' and 'Sales_Person' columns.")

                dedup_columns = [
                    col for col in headers
                    if module.ENABLE_RELATED_TAG or col not in {"TAG_NO", "TAG"}
                ] + ["NORMALIZED PARTY"]
                removed_filter_rows = int(state.get("removed_filter_rows", 0))
                duplicate_filter_rows = int(state.get("duplicate_filter_rows", 0))
                blank_filter_rows = int(state.get("blank_filter_rows", 0))
                processed_filter_rows = int(state.get("processed_filter_rows", 0))
                matched_filter_rows = int(state.get("matched_filter_rows", 0))
                matched_stock_links = int(state.get("matched_stock_links", 0))
                first_match_seq = int(state.get("first_match_seq", 0))
                pending_since_commit = 0
                rank_map = {"Exact": 0, "Nearby": 1, "Maybe": 2}
                # Reuse identical *matching predicates* across different parties or
                # metadata. Match semantics still come from the untouched master
                # engine; country, city, salesperson and related tags are applied
                # separately for each rule below. REMARKS must be part of the key
                # because INCLUDE and EXCLUDE have different matching semantics.
                predicate_columns = [
                    col for col in headers
                    if module.normalize_column_key(col) not in module.METADATA_FILTER_KEYS
                    or module.normalize_column_key(col) == "REMARKS"
                ]
                match_cache = OrderedDict()
                cache_row_count = 0
                cache_hits = int(state.get("cache_hits", 0))
                MAX_CACHE_RULES = 256
                MAX_CACHE_MATCHES = 20000

                print(f"Streaming Change Format in batches of {int(rule_batch_size):,} rules...")
                for raw_row in row_iter:
                    if consumed >= batch_rows or (consumed and time.monotonic() - batch_started >= max_batch_seconds):
                        stopped_for_checkpoint = True
                        break
                    consumed += 1
                    state["filter_offset"] += 1
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

                    predicate_key = _stable_rule_digest(rule, predicate_columns)
                    if predicate_key in match_cache:
                        cached_pairs = match_cache[predicate_key]
                        match_cache.move_to_end(predicate_key)
                        cache_hits += 1
                        pairs_source = cached_pairs
                        match_count = len(cached_pairs)
                    else:
                        candidate_labels = candidate_index.candidates_for_rule(rule, tolerance_level)
                        if candidate_labels is not None:
                            if len(candidate_labels) == 0:
                                match_types = None
                            else:
                                working_df = stocks_df.loc[candidate_labels]
                                working_cache = None
                        else:
                            working_df = stocks_df
                            if stock_cache is None:
                                stock_cache = module.build_stock_filter_cache(stocks_df)
                            working_cache = stock_cache

                        if candidate_labels is None or len(candidate_labels):
                            rule_series = pd.Series(rule)
                            matched = module.apply_tolerance_filter(
                                working_df, rule_series, tolerance_level, stock_cache=working_cache
                            )
                            match_types = (
                                module.classify_match_types_vectorized(matched, rule_series)
                                if not matched.empty else None
                            )
                        match_count = len(match_types) if match_types is not None else 0
                        if match_count <= MAX_CACHE_MATCHES:
                            cached_pairs = (
                                tuple((int(idx), rank_map.get(str(kind), 2)) for idx, kind in match_types.items())
                                if match_count else ()
                            )
                            while match_cache and (
                                len(match_cache) >= MAX_CACHE_RULES
                                or cache_row_count + match_count > MAX_CACHE_MATCHES
                            ):
                                _, evicted = match_cache.popitem(last=False)
                                cache_row_count -= len(evicted)
                            match_cache[predicate_key] = cached_pairs
                            cache_row_count += match_count
                            pairs_source = cached_pairs
                        else:
                            # Broad rules can match all stocks. Do not retain these
                            # giant result lists in the in-memory reuse cache.
                            pairs_source = (
                                (int(idx), rank_map.get(str(kind), 2))
                                for idx, kind in match_types.items()
                            )

                    if match_count == 0:
                        if pending_since_commit >= rule_batch_size:
                            print(f"  Processed {processed_filter_rows:,} unique rules...")
                            pending_since_commit = 0
                        continue

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

                    # sqlite3 accepts an iterator: do not create a second large
                    # list of every matched stone for broad rules.
                    match_rows = (
                        (
                            salesperson,
                            norm_party,
                            stock_idx,
                            match_rank,
                            tag_value,
                        )
                        for stock_idx, match_rank in pairs_source
                    )
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
                    matched_stock_links += match_count

                    if pending_since_commit >= rule_batch_size:
                        print(
                            f"  Processed {processed_filter_rows:,} unique rules; "
                            f"{matched_filter_rows:,} produced matches..."
                        )
                        pending_since_commit = 0
                # An exhausted iterator or a worksheet boundary marks the last filter batch.
                if not stopped_for_checkpoint:
                    state["phase"] = "reporting"
            finally:
                filter_wb.close()
            for key in ("removed_filter_rows", "duplicate_filter_rows", "blank_filter_rows",
                        "processed_filter_rows", "matched_filter_rows", "matched_stock_links",
                        "first_match_seq", "cache_hits"):
                state[key] = int(locals()[key])


            _save_state(conn, state)
            return state
        if state["phase"] == "reporting":
            temp_dir = job_dir
            report_limit = state["report_index"] + max_salespeople
            stocks_df = _load_normalized_stock(job_dir, file_path_stock, module)

            exclusion_data = pd.DataFrame()
            if module.ENABLE_CLEAN_FILE and clean_file_path and os.path.exists(clean_file_path):
                exclusion_data = module.load_exclusion_data(clean_file_path)
                print(f"No Need rules loaded: {len(exclusion_data):,}")
            elif module.ENABLE_CLEAN_FILE:
                print("No Need workbook not supplied: clean/exclusion rules skipped.")
            else:
                print("Clean/exclusion rules disabled by configuration.")
            global_exclusion_by_sp, specific_exclusion_by_sp_party = module.build_exclusion_rule_index(exclusion_data)

            reports_created = int(state.get("reports_created", 0))
            separated_party_files = int(state.get("separated_party_files", 0))
            below_minimum_parties = int(state.get("below_minimum_parties", 0))
            used_salesperson_filenames = set(state.get("used_salesperson_filenames", []))
            print(f"Building salesperson report ZIP (layout: {output_layout})...")
            if min_stones_per_party:
                print(f"Minimum stones per party: {min_stones_per_party} (checked after output filters)")
            with _StagedZipWriter(job_dir) as zipf:
                salesperson_rows = conn.execute(
                    "SELECT salesperson FROM salesperson_meta ORDER BY first_seq"
                ).fetchall()
                for report_index, (salesperson,) in enumerate(salesperson_rows):
                    if report_index < state["report_index"]:
                        continue
                    if report_index >= report_limit:
                        break
                    before_members = len(zipf.members)
                    # A salesperson workbook can contain many sheets; normal
                    # openpyxl holds every styled cell in memory until .save().
                    # Use its disk-backed write-only writer once the group is large.
                    total_group_links = (
                        conn.execute("SELECT COUNT(*) FROM matches WHERE salesperson = ?", (salesperson,)).fetchone()[0]
                        if include_workbooks else 0
                    )
                    stream_report = True  # bound memory for all checkpoint reports
                    workbook = Workbook(write_only=stream_report) if include_workbooks else None
                    nav_stream = None
                    if workbook is not None:
                        if stream_report:
                            nav_stream = workbook.create_sheet(title="Navigation")
                        else:
                            workbook.remove(workbook.active)
                    sheets_for_nav = []
                    global_exclusion_rules = global_exclusion_by_sp.get(salesperson, pd.DataFrame())
                    party_rows = conn.execute(
                        "SELECT norm_party, display_name, country, city FROM party_meta"
                        " WHERE salesperson = ? ORDER BY first_seq",
                        (salesperson,),
                    ).fetchall()
                    # Navigation must not collide with any party sheet.
                    used_sheet_names = {"navigation"}
                    used_party_filenames = set()
                    clean_sp = "".join(
                        char for char in str(salesperson) if char.isalnum() or char in (" ", "_")
                    ).strip()[:50] or "Report"
                    unique_sp = clean_sp
                    suffix = 1
                    while unique_sp.casefold() in used_salesperson_filenames:
                        extra = f"_{suffix}"
                        unique_sp = f"{clean_sp[:50 - len(extra)]}{extra}"
                        suffix += 1
                    used_salesperson_filenames.add(unique_sp.casefold())
                    nested_zip_path = os.path.join(temp_dir, f"{unique_sp}_parties.zip")

                    # Each salesperson gets a nested ZIP only when that layout is enabled.
                    nested_context = (
                        zipfile.ZipFile(nested_zip_path, "w", zipfile.ZIP_DEFLATED)
                        if include_separate else nullcontext(None)
                    )
                    with nested_context as party_zip:
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
                            # Apply the limit to the final unique stone rows, not rule count
                            # or the raw pre-exclusion matches; it is per salesperson/party.
                            if len(result) < min_stones_per_party:
                                below_minimum_parties += 1
                                continue

                            if "Sr.No" in result.columns:
                                result.drop(columns=["Sr.No"], inplace=True)
                            result.insert(0, "Sr.No", range(1, 1 + len(result)))
                            result = _order_output_columns(module, result)

                            valid_sheet_name = re.sub(r'[\\/*?:"<>|]', "_", str(party_display_name))[:31].strip() or "PartySheet"
                            original_name = valid_sheet_name
                            counter = 1
                            while valid_sheet_name.casefold() in used_sheet_names:
                                extra = f"_{counter}"
                                valid_sheet_name = f"{original_name[:31 - len(extra)]}{extra}"
                                counter += 1
                            used_sheet_names.add(valid_sheet_name.casefold())

                            if workbook is not None:
                                if stream_report:
                                    _add_styled_party_sheet_streaming(module, workbook, valid_sheet_name, result)
                                else:
                                    _add_styled_party_sheet(module, workbook, valid_sheet_name, result)
                            if party_zip is not None:
                                # Independent Excel files contain only their own party sheet,
                                # since a Navigation link to absent sheets would be invalid.
                                stream_single = True  # bound memory for all checkpoint reports
                                single_book = Workbook(write_only=stream_single)
                                if stream_single:
                                    _add_styled_party_sheet_streaming(module, single_book, valid_sheet_name, result)
                                else:
                                    single_book.remove(single_book.active)
                                    _add_styled_party_sheet(module, single_book, valid_sheet_name, result)
                                temp_xlsx = os.path.join(temp_dir, f"advanced_separate_{unique_sp}.xlsx")
                                try:
                                    single_book.save(temp_xlsx)
                                finally:
                                    single_book.close()
                                party_zip.write(
                                    temp_xlsx,
                                    arcname=_unique_party_filename(party_display_name, used_party_filenames),
                                )
                                os.remove(temp_xlsx)
                                separated_party_files += 1
                            sheets_for_nav.append(
                                (valid_sheet_name, party_display_name, len(result), country, city)
                            )

                    if not sheets_for_nav:
                        if workbook is not None:
                            workbook.close()
                        if include_separate:
                            os.remove(nested_zip_path)
                        state["report_index"] += 1
                        state["used_salesperson_filenames"] = list(used_salesperson_filenames)
                        state["below_minimum_parties"] = below_minimum_parties
                        _save_state(conn, state)
                        continue

                    if workbook is not None:
                        if stream_report:
                            _add_navigation_sheet_streaming(workbook, nav_stream, sheets_for_nav)
                        else:
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
                                escaped_name = sheet_name.replace("'", "''")
                                nav_ws.cell(row=row_num, column=6, value="Go \u2192").hyperlink = f"#'{escaped_name}'!A1"
                            module.apply_styles_to_worksheet(nav_ws)
                            module.auto_adjust_column_widths(nav_ws)
                        temp_xlsx = os.path.join(temp_dir, f"{unique_sp}.xlsx")
                        try:
                            workbook.save(temp_xlsx)
                        finally:
                            workbook.close()
                        zipf.write(temp_xlsx, arcname=f"{unique_sp}.xlsx")
                        os.remove(temp_xlsx)
                    if include_separate:
                        zipf.write(nested_zip_path, arcname=f"{unique_sp}.zip")
                        os.remove(nested_zip_path)
                    reports_created += 1
                    state["report_index"] += 1
                    state["used_salesperson_filenames"] = list(used_salesperson_filenames)
                    state["members"].extend(zipf.members[before_members:])
                    state["reports_created"] = reports_created
                    state["separated_party_files"] = separated_party_files
                    state["below_minimum_parties"] = below_minimum_parties
                    _save_state(conn, state)


            if state["report_index"] >= len(salesperson_rows):
                state["phase"] = "packaging"
                _save_state(conn, state)

            return state
        raise RuntimeError(f"Unknown processing phase: {state['phase']}")
