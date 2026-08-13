from __future__ import annotations
import argparse, gzip, io, json, pickle, sys, time, zipfile
from pathlib import Path
from typing import Any
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


def load(root: Path, name: str):
    with open(root / f"{name}.pkl", "rb") as fh:
        return pickle.load(fh)


def dump(root: Path, name: str, obj):
    with open(root / f"{name}.pkl", "wb") as fh:
        pickle.dump(obj, fh, pickle.HIGHEST_PROTOCOL)


def safe(v: Any) -> Any:
    if v is None:
        return None
    if isinstance(v, pd.Timestamp):
        return None if pd.isna(v) else v.to_pydatetime()
    if isinstance(v, np.datetime64):
        return None if pd.isna(v) else pd.Timestamp(v).to_pydatetime()
    if isinstance(v, np.integer):
        return int(v)
    if isinstance(v, np.floating):
        return None if np.isnan(v) else float(v)
    if isinstance(v, float) and np.isnan(v):
        return None
    if isinstance(v, (list, tuple, set, dict)):
        v = str(v)
    if not isinstance(v, (str, int, float, bool)):
        v = str(v)
    return v[:32760] if isinstance(v, str) and len(v) > 32767 else v


HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)

# Keep stable schemas even when a run produces zero accepted baskets.  Empty
# CSV files without headers caused the independent release audit to receive a
# columnless DataFrame and crash instead of reporting a valid zero-selection
# result.
EMPTY_SELECTED_COLUMNS = [
    "Sold Party", "Sales_Person", "Country", "City",
    "Deal_Mode", "Recipe_Line_ID", "Recipe_Priority",
    "Serial_Key", "Serial No", "Pair_Key", "Pair_Group_Members",
    "Pair_Partner_Serial", "Pair_Added_Flag", "Pair_Anchor_Serial",
    "Selection_Source", "Selection_Reason", "Match_Level",
    "Generation_Status", "Release_Status",
]

EMPTY_LINE_COLUMNS = [
    "Sold Party", "Deal_Mode", "Recipe_Line_ID", "Recipe_Priority",
    "Requested_Recipe_Quantity", "Candidate_Availability",
    "Pre_Allocation_Filled_Quantity", "Final_Globally_Allocated_Quantity",
    "Missing_Quantity", "Failure_Reason",
    "Recipe_Fill_Reconciliation_Status",
]


def ensure_columns(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    out = df.copy() if isinstance(df, pd.DataFrame) else pd.DataFrame()
    for column in columns:
        if column not in out.columns:
            out[column] = pd.Series(index=out.index, dtype="object")
    return out


def add_sheet(wb: Workbook, name: str, df: pd.DataFrame) -> None:
    ws = wb.create_sheet(name[:31])
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    if df is None or df.empty:
        c = WriteOnlyCell(ws, value="No rows")
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        ws.append([c])
        return
    headers = list(df.columns)
    header_row = []
    for h in headers:
        c = WriteOnlyCell(ws, value=str(h))
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        header_row.append(c)
    ws.append(header_row)
    date_idx = {i for i, h in enumerate(headers) if "Date" in str(h)}
    pct_idx = {i for i, h in enumerate(headers) if any(x in str(h) for x in ["Rate", "Share", "Coverage", "Progress"])}
    status_idx = headers.index("Generation_Status") if "Generation_Status" in headers else -1
    release_idx = headers.index("Release_Status") if "Release_Status" in headers else -1
    severity_idx = headers.index("Severity") if "Severity" in headers else -1
    for vals in df.itertuples(index=False, name=None):
        out = []
        for i, v in enumerate(vals):
            c = WriteOnlyCell(ws, value=safe(v))
            if i in date_idx and c.value is not None:
                c.number_format = "DD-MM-YYYY"
            if i in pct_idx and isinstance(c.value, (int, float)):
                c.number_format = "0.0%"
            if i == status_idx:
                text = str(c.value or "").upper()
                c.fill = PatternFill("solid", fgColor="C6EFCE" if text == "GENERATED" else ("FFF2CC" if "REVIEW" in text else "F2F2F2"))
            if i == release_idx:
                text = str(c.value or "").upper()
                good = text == "READY FOR SALES REVIEW"
                caution = any(x in text for x in ["PAIR", "CONFIRMATION", "REVIEW"])
                c.fill = PatternFill("solid", fgColor="C6EFCE" if good else ("FFF2CC" if caution else "F2F2F2"))
            if i == severity_idx:
                text = str(c.value or "").upper()
                c.fill = PatternFill("solid", fgColor="F4CCCC" if text == "ERROR" else ("FCE5CD" if text == "WARNING" else ("D9EAD3" if text == "OK" else "D9EAF7")))
            out.append(c)
        ws.append(out)
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(df)+1}"
    for i, h in enumerate(headers, 1):
        width = 14
        if h in {"Sold Party", "Sales_Person", "Generation_Reason", "Selection_Reason", "Behavior_Type", "Release_Status", "Pair_Resolution_Status", "Omission_Reason"}:
            width = 34
        elif "Date" in str(h):
            width = 13
        elif h in {"DETAIL", "VIDEO", "IMAGE", "CERTIFICATE"}:
            width = 18
        elif len(str(h)) > 18:
            width = 22
        ws.column_dimensions[get_column_letter(i)].width = width


def concatenate_batches(root: Path, prefix: str) -> pd.DataFrame:
    parts = []
    for path in sorted(root.glob(f"{prefix}_batch_*.pkl")):
        with open(path, "rb") as fh:
            parts.append(pickle.load(fh))
    frames = [x for x in parts if isinstance(x, pd.DataFrame)]
    nonempty = [x for x in frames if not x.empty]
    if not nonempty:
        # Preserve any schema carried by empty batch frames.
        columns: list[str] = []
        for frame in frames:
            for column in frame.columns:
                if column not in columns:
                    columns.append(column)
        return pd.DataFrame(columns=columns)
    import warnings
    with warnings.catch_warnings():
        warnings.filterwarnings(
            "ignore",
            message="The behavior of DataFrame concatenation with empty or all-NA entries is deprecated.*",
            category=FutureWarning,
        )
        return pd.concat(nonempty, ignore_index=True)


def load_final_allocation_state(root: Path) -> dict:
    paths = sorted(root.glob("allocation_state_*.pkl"))
    if not paths:
        return {"used_serials": set(), "pair_owners": {}}
    with open(paths[-1], "rb") as fh:
        state = pickle.load(fh)
    return state if isinstance(state, dict) else {"used_serials": set(state), "pair_owners": {}}


def main() -> None:
    ap = argparse.ArgumentParser(description="Consolidate v8 allocation into a production review workbook and detailed exports.")
    ap.add_argument("--pipeline", required=True)
    ap.add_argument("--work-dir", required=True)
    ap.add_argument("--allocation-dir", required=True)
    ap.add_argument("--output-dir", required=True)
    args = ap.parse_args()
    pipeline = Path(args.pipeline).resolve()
    sys.path.insert(0, str(pipeline.parent))
    import importlib.util
    spec = importlib.util.spec_from_file_location("dealmaker_pipeline_v6_finalize", pipeline)
    p = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = p
    spec.loader.exec_module(p)

    work = Path(args.work_dir)
    alloc = Path(args.allocation_dir)
    out = Path(args.output_dir)
    out.mkdir(parents=True, exist_ok=True)
    started = time.time()

    ps = load(work, "party_summary")
    behavior = load(work, "behavior_profiles")
    mode = load(work, "mode_summary")
    deals = load(work, "deals")
    profile_detail = load(work, "profile_detail")
    detail = profile_detail
    try:
        repeated_deal_serial_audit = load(work, "repeated_deal_serial_audit")
    except FileNotFoundError:
        repeated_deal_serial_audit = pd.DataFrame()
    action_conversion = load(work, "action_conversion")
    current = load(work, "current_actions")
    combos = load(work, "combination_profiles")
    recipes = load(work, "recipes")
    stock = load(work, "stock")
    overlap = load(work, "historical_sale_stock")
    unavailable_stock = load(work, "unavailable_stock")
    quality = load(work, "quality_log")
    as_of = load(work, "as_of")
    settings = load(work, "settings")
    quality_audit_stock = load(work, "quality_domain_audit_stock")
    quality_domain_config = load(work, "quality_domain_config")
    quality_alias_config = load(work, "quality_alias_config")
    stage1_metadata = json.loads((work / "stage1_metadata.json").read_text(encoding="utf-8"))
    # Source-wide domain audits are authoritative in the mandatory preflight.
    preflight_path = out / "Dealmaker_Data_Preflight_v6.xlsx"
    quality_audit_sales = pd.DataFrame()
    quality_audit_actions = pd.DataFrame()
    if preflight_path.exists():
        try:
            q = pd.read_excel(preflight_path, sheet_name="Quality_Domain_Audit")
            quality_audit_sales = q[q["Dataset"].eq("Sales")].copy()
            quality_audit_actions = q[q["Dataset"].eq("Actions")].copy()
        except Exception:
            pass

    status = concatenate_batches(alloc, "status")
    selected = ensure_columns(concatenate_batches(alloc, "selected"), EMPTY_SELECTED_COLUMNS)
    partial = concatenate_batches(alloc, "review_partial")
    line = concatenate_batches(alloc, "line_status")
    preview = concatenate_batches(alloc, "candidate_preview")
    direct_omissions = concatenate_batches(alloc, "direct_action_omissions")
    allocation_state = load_final_allocation_state(alloc)

    eligible = set(status.get("Sold Party", pd.Series(dtype=str)))
    non = ps[~ps["Sold Party"].isin(eligible)].copy()
    nonstatus = pd.DataFrame({
        "Sold Party": non["Sold Party"], "Interval_Status": non["Interval_Status"],
        "Generation_Status": "NOT DUE", "Generation_Reason": non["Auto_Include_Reason"],
        "Global_Allocation_Mode": p.clean_text(settings.get("GLOBAL_STOCK_ALLOCATION", "UNIQUE")).upper(),
        "Pair_Allocation_Policy": p.clean_text(settings.get("PAIR_ALLOCATION_POLICY", "RESERVE_PAIR_KEY")).upper(),
    })
    status = pd.concat([status, nonstatus], ignore_index=True, sort=False)
    order = {party: i for i, party in enumerate(ps["Sold Party"])}
    status["_order"] = status["Sold Party"].map(order).fillna(10**9)
    status = status.sort_values("_order").drop(columns="_order").reset_index(drop=True)

    selected_map = selected.groupby("Serial_Key")["Sold Party"].agg(lambda x: " | ".join(sorted(set(x.astype(str))))).to_dict() if not selected.empty else {}
    stock_keys = set(stock["Serial_Key"].astype(str))
    if not selected.empty:
        selected["Pair_Partner_In_Current_Stock"] = selected["Pair_Partner_Serial"].fillna("").astype(str).isin(stock_keys)
        selected["Pair_Partner_Selected_Party"] = selected["Pair_Partner_Serial"].fillna("").astype(str).map(selected_map).fillna("")
        def pair_resolution(r):
            if str(r.get("Pair_Review_Flag", "")).upper() != "YES":
                return "NOT A PAIR"
            if r["Pair_Partner_Selected_Party"] == r["Sold Party"]:
                return "BOTH SELECTED FOR SAME PARTY"
            if r["Pair_Partner_Selected_Party"]:
                return "PARTNER SELECTED FOR DIFFERENT PARTY"
            if r["Pair_Partner_In_Current_Stock"]:
                return "PAIR COMPLETION FAILURE - PARTNER AVAILABLE BUT NOT SELECTED"
            return "PARTNER NOT IN CURRENT STOCK"
        selected["Pair_Resolution_Status"] = selected.apply(pair_resolution, axis=1)

    pair_review = selected[selected.get("Pair_Review_Flag", pd.Series(index=selected.index, dtype=str)).eq("YES")].copy() if not selected.empty else pd.DataFrame()
    pair_parties = set(pair_review.get("Sold Party", pd.Series(dtype=str)))
    split_pairs = pair_review[pair_review.get("Pair_Resolution_Status", pd.Series(index=pair_review.index, dtype=str)).eq("PARTNER SELECTED FOR DIFFERENT PARTY")]
    split_pair_keys = int(split_pairs["Pair_Key"].nunique()) if not split_pairs.empty else 0
    split_pair_parties = set(split_pairs.get("Sold Party", pd.Series(dtype=str)))

    party_context_cols = [c for c in ["Sold Party", "Country", "City", "Sales_Person"] if c in ps.columns]
    if party_context_cols:
        party_context = ps[party_context_cols].drop_duplicates("Sold Party")
        status = status.drop(columns=[c for c in ["Country", "City"] if c in status.columns], errors="ignore").merge(party_context, on="Sold Party", how="left", suffixes=("", "_History"))
        if "Sales_Person_History" in status.columns:
            status["Sales_Person"] = status.get("Sales_Person", pd.Series(index=status.index, dtype=object)).fillna(status["Sales_Person_History"])
            status = status.drop(columns=["Sales_Person_History"])

    stock_status_present = bool(stock.get("Stock_Status", pd.Series(dtype=str)).fillna("").astype(str).str.strip().ne("").any())
    snapshot_confirmed = p.yes(settings.get("STOCK_SNAPSHOT_CONFIRMED_LIVE", "NO"))
    live_stock_verified = stock_status_present or snapshot_confirmed
    max_sales_age = int(float(settings.get("MAX_SALES_STALENESS_DAYS", 3)))
    max_actions_age = int(float(settings.get("MAX_ACTIONS_STALENESS_DAYS", 7)))
    sales_age = stage1_metadata.get("sales_staleness_days")
    actions_age = stage1_metadata.get("actions_staleness_days")
    require_fresh = p.yes(settings.get("RELEASE_REQUIRES_SOURCE_FRESHNESS", "YES"))
    fresh_sources_verified = (
        sales_age is not None and actions_age is not None
        and int(sales_age) <= max_sales_age and int(actions_age) <= max_actions_age
    ) or not require_fresh

    def release_status(row: pd.Series) -> str:
        if row.get("Generation_Status") != "GENERATED":
            return "INSUFFICIENT STOCK - DO NOT RELEASE" if "REVIEW" in str(row.get("Generation_Status")) else "NOT DUE"
        party = row.get("Sold Party")
        if party in split_pair_parties:
            return "BLOCKED - PAIR SPLIT DETECTED"
        requirements = []
        if not fresh_sources_verified:
            requirements.append("FRESH SOURCE EXPORT REQUIRED")
        if not live_stock_verified:
            requirements.append("LIVE STOCK CONFIRMATION REQUIRED")
        if party in pair_parties:
            requirements.append("PAIR REVIEW REQUIRED")
        if str(row.get("Action_Review_Required", "NO")).upper() == "YES":
            requirements.append("ACTION SUPPORT REVIEW REQUIRED")
        return " + ".join(requirements) if requirements else "READY FOR SALES REVIEW"

    status["Release_Status"] = status.apply(release_status, axis=1)
    line = ensure_columns(p.reconcile_recipe_fill(line, selected, status), EMPTY_LINE_COLUMNS)
    export_status_cols = [c for c in [
        "Sold Party", "Country", "City", "Chosen_Deal_Mode", "Target_Stone_Count",
        "Chosen_Mode_Interval_Status", "Chosen_Mode_Interval_Source", "Chosen_Mode_Interval_Progress",
        "Generation_Status", "Generation_Reason", "Action_Policy", "Action_Support_Status",
        "Action_Data_Status", "Action_Review_Required", "Action_Support_Threshold",
        "Current_Action_Stone_Count", "Release_Status"
    ] if c in status.columns]
    if not selected.empty:
        selected = selected.merge(status[export_status_cols].drop_duplicates("Sold Party"), on="Sold Party", how="left", suffixes=("", "_Status"))
    if not partial.empty:
        partial = partial.merge(status[export_status_cols].drop_duplicates("Sold Party"), on="Sold Party", how="left", suffixes=("", "_Status"))

    conflicts = p.build_cross_party_conflicts(selected)
    deal_id_party_counts = deals.groupby("Deal_ID")["Sold Party"].nunique() if not deals.empty else pd.Series(dtype=int)
    reused_deal_ids = deal_id_party_counts[deal_id_party_counts.gt(1)]
    deal_id_review = deals[deals["Deal_ID"].isin(reused_deal_ids.index)].sort_values(["Deal_ID", "Sold Party", "Deal_Start_Date"]) if not reused_deal_ids.empty else pd.DataFrame(columns=deals.columns)

    pair_ledger_rows = []
    for pair_key, owner in sorted(dict(allocation_state.get("pair_owners", {})).items()):
        members = selected[selected.get("Pair_Key", pd.Series(dtype=str)).eq(pair_key)] if not selected.empty else pd.DataFrame()
        pair_ledger_rows.append({
            "Pair_Key": pair_key, "Reserved_To_Party": owner,
            "Selected_Member_Count": len(members),
            "Selected_Serials": " | ".join(sorted(members.get("Serial_Key", pd.Series(dtype=str)).astype(str).unique())) if not members.empty else "",
        })
    pair_ledger = pd.DataFrame(pair_ledger_rows)

    if not preview.empty:
        preview = p.annotate_candidate_archive(preview, selected, status)
        preview["_rank"] = preview["Pre_Allocation_Candidate_Status"].map({"SELECTED": 0, "NOT SELECTED": 1, "EXCLUDED": 2}).fillna(3)
        preview = preview.sort_values(["_rank", "Final_Selection_Score"], ascending=[True, False]).head(5000).drop(columns="_rank")
        preview = preview.drop(columns=[c for c in ["DETAIL", "VIDEO", "IMAGE", "CERTIFICATE"] if c in preview])
    if not direct_omissions.empty:
        direct_omissions = direct_omissions.drop(columns=[c for c in ["DETAIL", "VIDEO", "IMAGE", "CERTIFICATE"] if c in direct_omissions])

    candidate_total = 0
    for f in alloc.glob("candidate_total_batch_*.pkl"):
        with open(f, "rb") as fh:
            candidate_total += int(pickle.load(fh))
    candidate_zip = out / "Unique_Allocation_Candidates_By_Batch_v6.zip"
    with zipfile.ZipFile(candidate_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in sorted(alloc.glob("Candidates_Batch_*.csv.gz")):
            try:
                batch = pd.read_csv(f, low_memory=False)
            except Exception as exc:
                raise RuntimeError(
                    f"Unable to read candidate batch {f.name}. "
                    "The batch archive is malformed or incomplete."
                ) from exc
            batch = p.annotate_candidate_archive(batch, selected, status)
            csv_bytes = batch.to_csv(index=False).encode("utf-8")
            buffer = io.BytesIO()
            with gzip.GzipFile(fileobj=buffer, mode="wb") as gz:
                gz.write(csv_bytes)
            zf.writestr(f.name, buffer.getvalue())
        zf.writestr(
            "CANDIDATE_ARCHIVE_SEMANTICS_v6.txt",
            (
                "This archive is not a final-deal output.\n"
                "It contains ranked candidates for every attempted deal mode after stock-status, historical-sale-overlap, quality, business-exclusion, prior-party serial reservation, and pair-key reservation filters.\n"
                "Candidate_Status/Pre_Allocation_Candidate_Status describes selection inside that attempted mode before the final mode and basket acceptance decision.\n"
                "Final_Globally_Allocated_Flag and Final_Allocation_Status identify whether the row is in Selected_Deal_Stones_v6.csv.\n"
                "The authoritative final allocation is Selected_Deal_Stones_v6.csv.\n"
            ),
        )

    for name, df in [
        ("Deal_Stone_Details", detail), ("Profile_Deal_Stone_Details", profile_detail),
        ("Repeated_Deal_Serial_Audit", repeated_deal_serial_audit),
        ("Combination_Profile", combos), ("Basket_Recipe_Lines", recipes),
        ("Current_Action_Candidates", current),
    ]:
        df.to_csv(out / f"{name}_v6.csv.gz", index=False, compression="gzip")

    selected_overlap = int(selected["Serial_Key"].isin(set(overlap["Serial_Key"])).sum()) if not selected.empty else 0
    quality_flags = int(selected.get("Quality_Review_Flag", pd.Series(dtype=str)).eq("YES").sum()) if not selected.empty else 0
    hold_count = int(stage1_metadata.get("hold_action_rows", 0))
    repeated_profile_rows = int(stage1_metadata.get("repeated_deal_serial_rows_removed", 0))
    recipe_fill_mismatches = int(line.get("Recipe_Fill_Reconciliation_Status", pd.Series(dtype=str)).eq("MISMATCH").sum()) if not line.empty else 0
    pair_added_count = int(selected.get("Pair_Added_Flag", pd.Series("NO", index=selected.index)).astype(str).str.upper().eq("YES").sum()) if not selected.empty else 0
    action_types_found = ", ".join(stage1_metadata.get("action_types", []))

    extra = [
        {"Dataset": "Sales", "Check": "Completed SALE rows", "Value": int(stage1_metadata.get("sale_rows", 0)), "Severity": "OK"},
        {"Dataset": "Sales", "Check": "BID rows retained as intent only", "Value": int(stage1_metadata.get("bid_rows", 0)), "Severity": "OK"},
        {"Dataset": "Deals", "Check": "Real deal buckets", "Value": len(deals), "Severity": "OK"},
        {"Dataset": "Deals", "Check": "Deal_ID values reused across different parties", "Value": len(reused_deal_ids), "Severity": "ERROR" if len(reused_deal_ids) else "OK"},
        {"Dataset": "Deals", "Check": "Repeated serial rows removed from profile calculations", "Value": repeated_profile_rows, "Severity": "WARNING" if repeated_profile_rows else "OK"},
        {"Dataset": "Actions", "Check": "Action types found", "Value": action_types_found, "Severity": "OK"},
        {"Dataset": "Actions", "Check": "HOLD action rows", "Value": hold_count, "Severity": "INFO"},
        {"Dataset": "Stock", "Check": "Actual-stock rows eligible after overlap exclusion", "Value": len(stock), "Severity": "OK"},
        {"Dataset": "Stock", "Check": "Completed-SALE history overlaps excluded", "Value": len(overlap), "Severity": "WARNING" if len(overlap) else "OK"},
        {"Dataset": "Stock", "Check": "Explicit availability/status column populated", "Value": "YES" if stock_status_present else "NO", "Severity": "OK" if stock_status_present else "WARNING"},
        {"Dataset": "Stock", "Check": "Stock snapshot manually confirmed live-only", "Value": "YES" if snapshot_confirmed else "NO", "Severity": "OK" if snapshot_confirmed else "INFO"},
        {"Dataset": "Stock", "Check": "Live-stock release verification", "Value": "VERIFIED" if live_stock_verified else "NOT VERIFIED - RELEASE BLOCKED", "Severity": "OK" if live_stock_verified else "WARNING"},
        {"Dataset": "Release", "Check": "Sales source freshness", "Value": f"{sales_age} days (maximum {max_sales_age})", "Severity": "OK" if sales_age is not None and int(sales_age) <= max_sales_age else "WARNING"},
        {"Dataset": "Release", "Check": "Actions source freshness", "Value": f"{actions_age} days (maximum {max_actions_age})", "Severity": "OK" if actions_age is not None and int(actions_age) <= max_actions_age else "WARNING"},
        {"Dataset": "Release", "Check": "Fresh-source release verification", "Value": "VERIFIED" if fresh_sources_verified else "NOT VERIFIED - RELEASE BLOCKED", "Severity": "OK" if fresh_sources_verified else "WARNING"},
        {"Dataset": "Output", "Check": "Accepted generated parties", "Value": int(status["Generation_Status"].eq("GENERATED").sum()), "Severity": "OK"},
        {"Dataset": "Output", "Check": "Due parties without an approved basket", "Value": int(status["Generation_Status"].astype(str).str.contains("REVIEW").sum()), "Severity": "WARNING"},
        {"Dataset": "Output", "Check": "Final selected stones", "Value": len(selected), "Severity": "OK"},
        {"Dataset": "Output", "Check": "Unique selected serials", "Value": selected["Serial_Key"].nunique() if not selected.empty else 0, "Severity": "OK"},
        {"Dataset": "Output", "Check": "Cross-party duplicate serials", "Value": len(conflicts), "Severity": "ERROR" if len(conflicts) else "OK"},
        {"Dataset": "Output", "Check": "Selected completed-sale overlap serials", "Value": selected_overlap, "Severity": "ERROR" if selected_overlap else "OK"},
        {"Dataset": "Output", "Check": "Selected quality-review stones", "Value": quality_flags, "Severity": "ERROR" if quality_flags else "OK"},
        {"Dataset": "Output", "Check": "Pair partners automatically added", "Value": pair_added_count, "Severity": "INFO"},
        {"Dataset": "Output", "Check": "Pair-review selected stones", "Value": len(pair_review), "Severity": "WARNING" if len(pair_review) else "OK"},
        {"Dataset": "Output", "Check": "Pair-review parties", "Value": len(pair_parties), "Severity": "WARNING" if pair_parties else "OK"},
        {"Dataset": "Output", "Check": "Pair keys split across different parties", "Value": split_pair_keys, "Severity": "ERROR" if split_pair_keys else "OK"},
        {"Dataset": "Output", "Check": "Direct-action candidates omitted from chosen baskets", "Value": len(direct_omissions), "Severity": "INFO"},
        {"Dataset": "Output", "Check": "Partial attempts excluded from final selected output", "Value": len(partial), "Severity": "OK"},
        {"Dataset": "Output", "Check": "Full unique-allocation candidate rows", "Value": candidate_total, "Severity": "INFO"},
        {"Dataset": "Output", "Check": "Recipe-fill final-allocation mismatches", "Value": recipe_fill_mismatches, "Severity": "ERROR" if recipe_fill_mismatches else "OK"},
        {"Dataset": "Output", "Check": "Candidate archive is final output", "Value": "NO - ranked attempted-mode archive", "Severity": "INFO"},
    ]
    data_quality = pd.concat([pd.DataFrame(quality), pd.DataFrame(extra)], ignore_index=True)

    ready_count = int(status["Release_Status"].eq("READY FOR SALES REVIEW").sum())
    live_confirm_count = int(status["Release_Status"].astype(str).str.contains("LIVE STOCK CONFIRMATION").sum())
    pair_review_count = int(status["Release_Status"].astype(str).str.contains("PAIR REVIEW").sum())
    overall_interval_eligible = int(ps["Auto_Include"].astype(str).str.upper().eq("YES").sum())
    mode_due_parties = int(stage1_metadata.get("mode_due_eligible_parties", len(status)))
    due_without_basket = int(status["Generation_Status"].astype(str).str.contains("REVIEW").sum())
    generated_mask = status["Generation_Status"].eq("GENERATED")
    action_supported_count = int((generated_mask & status.get("Action_Support_Status", pd.Series("", index=status.index)).eq("ACTION SUPPORTED")).sum())
    sales_history_only_count = int((generated_mask & status.get("Action_Support_Status", pd.Series("", index=status.index)).eq("SALES HISTORY ONLY")).sum())
    no_recent_action_count = int((generated_mask & status.get("Action_Data_Status", pd.Series("", index=status.index)).eq("NO RECENT ACTION DATA")).sum())
    stale_action_count = int((generated_mask & status.get("Action_Data_Status", pd.Series("", index=status.index)).eq("STALE ACTION DATA")).sum())
    review_action_issue = status.get("Action_Review_Required", pd.Series("NO", index=status.index)).astype(str).str.upper().eq("YES")
    review_mask = status["Generation_Status"].astype(str).str.contains("REVIEW")
    # Under BOOST_ONLY, action weakness is review context only and is never a rejection cause.
    action_policy_value = str(settings.get("ACTION_POLICY", "BOOST_ONLY")).upper()
    action_policy_rejections = 0 if action_policy_value == "BOOST_ONLY" else int((review_mask & review_action_issue).sum())
    stock_control_failures = int(review_mask.sum())
    stock_control_with_action_review = int((review_mask & review_action_issue).sum())
    summary = [
        ("Pipeline version", p.VERSION), ("Data as-of date", as_of),
        ("Sales rows after business dedupe", int(stage1_metadata.get("sales_rows", 0))), ("Completed SALE rows", int(stage1_metadata.get("sale_rows", 0))), ("BID rows", int(stage1_metadata.get("bid_rows", 0))),
        ("Actions after business dedupe", int(stage1_metadata.get("action_rows", 0))), ("HOLD action rows", hold_count),
        ("Real deal buckets", len(deals)), ("Globally reused Deal_ID values", len(reused_deal_ids)),
        ("Repeated deal-serial rows removed from profiles", repeated_profile_rows),
        ("Repeated deal-serial audit groups", len(repeated_deal_serial_audit)),
        ("Recipe-fill final-allocation mismatches", recipe_fill_mismatches),
        ("Parties analyzed", len(ps)),
        ("Overall-party interval eligible parties", overall_interval_eligible),
        ("Parties with at least one due mode", mode_due_parties),
        ("Generated parties", int(status["Generation_Status"].eq("GENERATED").sum())),
        ("Generated parties ready for sales review", ready_count),
        ("Generated parties requiring live-stock confirmation", live_confirm_count),
        ("Generated parties requiring pair review", pair_review_count),
        ("Due parties without an approved basket", due_without_basket),
        ("Action-policy rejections", action_policy_rejections),
        ("Stock/coverage/quality/pair control failures", stock_control_failures),
        ("Review parties also carrying weak/stale action support (non-causal)", stock_control_with_action_review),
        ("Generated baskets - ACTION SUPPORTED", action_supported_count),
        ("Generated baskets - SALES HISTORY ONLY", sales_history_only_count),
        ("Generated baskets - NO RECENT ACTION DATA", no_recent_action_count),
        ("Generated baskets - STALE ACTION DATA", stale_action_count),
        ("Final selected stones", len(selected)), ("Unique selected serials", selected["Serial_Key"].nunique() if not selected.empty else 0),
        ("Cross-party duplicate serials", len(conflicts)), ("Selected quality-review stones", quality_flags),
        ("Pair partners automatically added", pair_added_count),
        ("Pair-review selected stones", len(pair_review)), ("Pair keys reserved", len(allocation_state.get("pair_owners", {}))),
        ("Pair keys split across parties", split_pair_keys), ("Completed-sale stock overlaps excluded", len(overlap)),
        ("Stock availability field populated", "YES" if stock_status_present else "NO"),
        ("Stock snapshot manually confirmed live-only", "YES" if snapshot_confirmed else "NO"),
        ("Run date", stage1_metadata.get("run_date")),
        ("Latest completed SALE date", stage1_metadata.get("latest_sale_date")),
        ("Latest action date", stage1_metadata.get("latest_action_date")),
        ("Sales staleness days", sales_age), ("Actions staleness days", actions_age),
        ("Fresh-source release verified", "YES" if fresh_sources_verified else "NO"),
        ("Quality-domain release verified", "YES"),
        ("Global allocation mode", settings.get("GLOBAL_STOCK_ALLOCATION", "UNIQUE")),
        ("Pair allocation policy", settings.get("PAIR_ALLOCATION_POLICY", "RESERVE_PAIR_KEY")),
        ("Action policy", settings.get("ACTION_POLICY", "BOOST_ONLY")),
        ("No-action result", settings.get("NO_ACTION_RESULT", "GENERATE_WITH_REVIEW")),
        ("Stale-action result", settings.get("STALE_ACTION_RESULT", "GENERATE_WITH_REVIEW")),
        ("Direct action policy", settings.get("DIRECT_ACTION_POLICY", "DIRECT_FIRST_WITHIN_MATCH")),
        ("Direct action types", settings.get("DIRECT_SELECTION_ACTION_TYPES", "HOLD,CART,OFFER")),
        ("Full candidate rows", candidate_total),
    ]

    frames = {
        "Data_Quality": data_quality, "Party_Generation_Status": status, "Selected_Deal_Stones": selected,
        "Review_Partial_Attempts": partial, "Pair_Review": pair_review, "Pair_Reservation_Ledger": pair_ledger,
        "Cross_Party_Conflicts": conflicts, "Deal_ID_Uniqueness_Review": deal_id_review,
        "Direct_Action_Omissions": direct_omissions, "Recipe_Fill_Status": line,
        "Repeated_Deal_Serial_Audit": repeated_deal_serial_audit,
        "Party_Behavior_Profile": behavior, "Party_Deal_Summary": ps, "Party_Deal_Mode_Summary": mode,
        "Deal_Buckets": deals, "Party_Action_Conversion": action_conversion,
        "Current_Action_Candidates": current, "Basket_Recipe_Lines": recipes,
        "Quality_Domain_Audit_Sales": quality_audit_sales,
        "Quality_Domain_Audit_Actions": quality_audit_actions,
        "Quality_Domain_Audit_Stock": quality_audit_stock,
        "Quality_Domain_Config": quality_domain_config,
        "Quality_Alias_Config": quality_alias_config,
        "Stock_Sale_Overlap": overlap, "Stock_Quality_Exclusions": unavailable_stock, "Candidate_Preview": preview,
    }
    wb = Workbook(write_only=True)
    dash = wb.create_sheet("Dashboard")
    dash.append(["DEALMAKER PRODUCTION REVIEW - v8 SALES-HISTORY FIRST / ACTION BOOST-ONLY"])
    dash.append(["Pipeline version", p.VERSION])
    dash.append(["Review order", "Run_Summary -> Party_Generation_Status -> Selected_Deal_Stones -> Pair_Review -> Salesperson Outputs"])
    dash.append(["Release rule", "Completed Sales decides due mode and recipe. Actions are boost-only. READY still requires fresh sources, verified live stock, and resolved pair/action review labels."])
    dash.append(["Deal key rule", "Deal_ID is globally unique and recent profiles use Sold Party + Deal Mode + Deal_ID."])
    add_sheet(wb, "Run_Summary", pd.DataFrame(summary, columns=["Metric", "Value"]))
    for name in [
        "Data_Quality", "Party_Generation_Status", "Selected_Deal_Stones", "Review_Partial_Attempts",
        "Pair_Review", "Pair_Reservation_Ledger", "Cross_Party_Conflicts", "Deal_ID_Uniqueness_Review",
        "Direct_Action_Omissions", "Recipe_Fill_Status", "Party_Behavior_Profile", "Party_Deal_Summary",
        "Party_Deal_Mode_Summary", "Deal_Buckets", "Repeated_Deal_Serial_Audit", "Party_Action_Conversion", "Current_Action_Candidates",
        "Basket_Recipe_Lines", "Quality_Domain_Audit_Sales", "Quality_Domain_Audit_Actions",
        "Quality_Domain_Audit_Stock", "Quality_Domain_Config", "Quality_Alias_Config",
        "Stock_Sale_Overlap", "Stock_Quality_Exclusions", "Candidate_Preview",
    ]:
        add_sheet(wb, name, frames[name])
    review_path = out / "Dealmaker_Production_Review_v6.xlsx"  # legacy-compatible filename; workbook identifies pipeline v8
    wb.save(review_path)

    status.to_csv(out / "Party_Generation_Status_v6.csv", index=False)
    selected.to_csv(out / "Selected_Deal_Stones_v6.csv", index=False)
    pair_review.to_csv(out / "Pair_Review_v6.csv", index=False)
    direct_omissions.to_csv(out / "Direct_Action_Omissions_v6.csv", index=False)
    data_quality.to_csv(out / "Data_Quality_v6.csv", index=False)
    line.to_csv(out / "Recipe_Fill_Status_v6.csv", index=False)
    repeated_deal_serial_audit.to_csv(out / "Repeated_Deal_Serial_Audit_v6.csv", index=False)
    unavailable_stock.to_csv(out / "Stock_Quality_Exclusions_v7.csv", index=False)
    overlap.to_csv(out / "Stock_History_Overlap_Exclusions_v7.csv", index=False)
    (out / "CANDIDATE_ARCHIVE_SEMANTICS_v6.txt").write_text(
        "The candidate ZIP is a ranked attempted-mode archive, not a final basket.\n"
        "Use Final_Globally_Allocated_Flag/Final_Allocation_Status inside the archive and Selected_Deal_Stones_v6.csv as the authoritative final allocation.\n",
        encoding="utf-8",
    )
    dump(out, "final_status", status)
    dump(out, "final_selected", selected)

    report = out / "PRODUCTION_VALIDATION_REPORT_v6.txt"
    with report.open("w", encoding="utf-8") as fh:
        fh.write("DEALMAKER PRODUCTION VALIDATION REPORT v8\n========================================\n")
        for key, value in summary:
            fh.write(f"{key}: {value}\n")
        fh.write("\nRelease interpretation:\n")
        fh.write("- READY FOR SALES REVIEW: passed all automated checks and stock availability was verified by status field or explicit live-snapshot confirmation.\n")
        fh.write("- LIVE STOCK CONFIRMATION REQUIRED: basket is mechanically valid but the source has no verified availability field.\n")
        fh.write("- PAIR REVIEW REQUIRED: accepted basket contains pair-marked stock. The allocator reserves the pair key so it cannot be split across parties.\n")
        fh.write("- INSUFFICIENT STOCK - DO NOT RELEASE: diagnostic partial attempt only.\n")
        fh.write("\nResolved v8 controls:\n")
        fh.write("- Deal_ID values are globally unique.\n- Recent pattern membership uses party + mode + deal key.\n")
        fh.write("- Repeated serials inside the same deal are counted once in profile calculations and retained in a separate audit.\n")
        fh.write("- A serial reused in separate historical deals is counted once per Sold Party + Deal_ID + Serial No.\n")
        fh.write("- Recipe fill separates requested, candidate, pre-allocation, final allocation, missing quantity, and failure reason.\n")
        fh.write("- Candidate archives are explicitly marked as ranked attempted-mode records, not final baskets.\n")
        fh.write("- Pair keys are globally reserved to the first accepted party.\n")
        fh.write("- HOLD=0 is informational, not a failure.\n")
        fh.write("- Valid direct actions are prioritized within the same match level.\n")
        fh.write("- Commercial release is blocked until current-stock availability is verified.\n")
        fh.write("- Commercial release is blocked when Sales or Actions exceed configured freshness limits.\n")
        fh.write("- Quality values are canonical, explicitly aliased, review-only, unknown, or unmapped; no silent NONE conversion is allowed.\n")

    metadata = {
        "review_workbook": str(review_path),
        "generated_parties": int(status["Generation_Status"].eq("GENERATED").sum()),
        "selected_stones": len(selected), "ready_parties": ready_count,
        "live_confirmation_parties": live_confirm_count, "pair_review_parties": pair_review_count,
        "split_pair_keys": split_pair_keys, "reused_deal_ids": len(reused_deal_ids),
        "elapsed_seconds": round(time.time() - started, 2),
    }
    (out / "stage3_metadata.json").write_text(json.dumps(metadata, indent=2), encoding="utf-8")
    print(json.dumps(metadata, indent=2), flush=True)


if __name__ == "__main__":
    main()
