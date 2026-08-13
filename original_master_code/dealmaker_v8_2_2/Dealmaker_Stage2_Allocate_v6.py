from __future__ import annotations
import argparse, json, os, pickle, sys, time
from pathlib import Path
import pandas as pd


def load(root: Path, name: str):
    with open(root / f"{name}.pkl", "rb") as fh:
        return pickle.load(fh)


def dump(path: Path, name: str, obj):
    with open(path / f"{name}.pkl", "wb") as fh:
        pickle.dump(obj, fh, pickle.HIGHEST_PROTOCOL)


def load_allocation_state(out: Path, chunk: int) -> dict:
    if chunk <= 0:
        return {"used_serials": set(), "pair_owners": {}}
    previous = load(out, f"allocation_state_{chunk-1:03d}")
    if isinstance(previous, set):  # defensive compatibility with v4 state
        return {"used_serials": previous, "pair_owners": {}}
    return {
        "used_serials": set(previous.get("used_serials", set())),
        "pair_owners": dict(previous.get("pair_owners", {})),
    }


def _concat_frames_no_future_warning(frames):
    """Concatenate batch frames while silencing a pandas dtype-only future warning.

    The warning is triggered by all-NA optional columns in valid candidate frames.
    Current and future pandas behavior does not affect row membership or business logic.
    """
    import warnings
    valid = [x for x in frames if isinstance(x, pd.DataFrame) and not x.empty]
    if not valid:
        return pd.DataFrame()
    with warnings.catch_warnings():
        warnings.filterwarnings(
            "ignore",
            message="The behavior of DataFrame concatenation with empty or all-NA entries is deprecated.*",
            category=FutureWarning,
        )
        return pd.concat(valid, ignore_index=True)


def write_candidate_batch(path: Path, frames) -> pd.DataFrame:
    """Write one candidate batch using a single union-column schema.

    Candidate modes may expose different optional audit columns. Writing each
    frame in append mode under the first frame's header creates malformed CSV
    rows. One consolidated write keeps every row aligned to the same schema.
    """
    candidate_df = _concat_frames_no_future_warning(frames)
    if candidate_df.empty:
        if path.exists():
            path.unlink()
        return candidate_df
    candidate_df.to_csv(
        path,
        index=False,
        compression="gzip",
        lineterminator="\n",
    )
    return candidate_df


def main() -> None:
    ap = argparse.ArgumentParser(description="Allocate stock to one sequential party batch while preserving serial and pair reservations.")
    ap.add_argument("--pipeline", required=True)
    ap.add_argument("--work-dir", required=True)
    ap.add_argument("--allocation-dir", required=True)
    ap.add_argument("--chunk", type=int, required=True)
    ap.add_argument("--chunk-size", type=int, default=10)
    args = ap.parse_args()

    pipeline = Path(args.pipeline).resolve()
    sys.path.insert(0, str(pipeline.parent))
    import importlib.util
    spec = importlib.util.spec_from_file_location("dealmaker_pipeline_v6_allocate", pipeline)
    p = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = p
    spec.loader.exec_module(p)

    root = Path(args.work_dir)
    out = Path(args.allocation_dir)
    out.mkdir(parents=True, exist_ok=True)
    ps = load(root, "party_summary")
    bp = load(root, "behavior_profiles")
    mode_summary = load(root, "mode_summary")
    recipes = load(root, "recipes")
    current_actions = load(root, "current_actions")
    stock = load(root, "stock")
    overrides = load(root, "overrides")
    removed = load(root, "removed")
    exclusions = load(root, "exclusions")
    settings = dict(load(root, "settings"))
    stage1_metadata = json.loads((root / "stage1_metadata.json").read_text()) if (root / "stage1_metadata.json").exists() else {}
    actions_staleness_days = stage1_metadata.get("actions_staleness_days")

    allocation = p.clean_text(settings.get("GLOBAL_STOCK_ALLOCATION", "UNIQUE")).upper()
    pair_policy = p.clean_text(settings.get("PAIR_ALLOCATION_POLICY", "RESERVE_PAIR_KEY")).upper()
    if pair_policy not in {"RESERVE_PAIR_KEY", "REVIEW_ONLY"}:
        raise ValueError("PAIR_ALLOCATION_POLICY must be RESERVE_PAIR_KEY or REVIEW_ONLY")

    removed_set = set(removed.get("Sold Party", pd.Series(dtype=str)).dropna().astype(str).str.strip()) if not removed.empty else set()
    behavior_map = bp.set_index("Sold Party").to_dict("index") if not bp.empty else {}
    stock_index = {
        (p.clean_text(k[0]).upper(), p.clean_text(k[1]), p.clean_text(k[2]).upper(), p.clean_text(k[3]).upper()): g
        for k, g in stock.groupby(["Shp", "Carat_grp", "Lab", "Clarity_Band"], sort=False, dropna=False)
    } if not stock.empty else {}
    action_index = {party: g for party, g in current_actions.groupby("Sold Party", sort=False)} if not current_actions.empty else {}

    include_due_soon = p.yes(settings.get("INCLUDE_DUE_SOON", "NO"))
    allowed_mode_statuses = {"DUE NOW", "OVERDUE"} | ({"DUE SOON"} if include_due_soon else set())
    if p.yes(settings.get("MODE_DUE_REQUIRED_FOR_SELECTION", "YES")):
        due_mode_rows = mode_summary[mode_summary["Mode_Interval_Status"].isin(allowed_mode_statuses)].copy()
        due_parties = set(due_mode_rows["Sold Party"].astype(str))
        eligible = ps[ps["Sold Party"].astype(str).isin(due_parties)].copy()
        max_progress = due_mode_rows.groupby("Sold Party")["Mode_Interval_Progress"].max()
        due_count = due_mode_rows.groupby("Sold Party")["Deal_Mode"].nunique()
        eligible["Mode_Interval_Progress_Max"] = eligible["Sold Party"].map(max_progress).fillna(-1)
        eligible["Due_Mode_Count"] = eligible["Sold Party"].map(due_count).fillna(0).astype(int)
        eligible = eligible.sort_values(
            ["Mode_Interval_Progress_Max", "Recent_Action_Count", "Deal_Bucket_Count", "Sold Party"],
            ascending=[False, False, False, True], kind="mergesort"
        ).reset_index(drop=True)
    else:
        eligible = ps[ps["Auto_Include"].astype(str).str.upper().eq("YES")].copy().sort_values(
            ["Interval_Progress_Ratio", "Recent_Action_Count", "Deal_Bucket_Count", "Sold Party"],
            ascending=[False, False, False, True], kind="mergesort"
        ).reset_index(drop=True)
    start = args.chunk * args.chunk_size
    end = min(start + args.chunk_size, len(eligible))
    subset = eligible.iloc[start:end]

    state = load_allocation_state(out, args.chunk) if allocation == "UNIQUE" else {"used_serials": set(), "pair_owners": {}}
    global_used = state["used_serials"]
    pair_owners = state["pair_owners"]

    status_rows = []
    selected_parts = []
    partial_parts = []
    line_parts = []
    preview_parts = []
    omission_parts = []
    candidate_parts = []
    candidate_total = 0
    candidate_path = out / f"Candidates_Batch_{args.chunk:03d}.csv.gz"
    started = time.time()

    for _, row in subset.iterrows():
        party = row["Sold Party"]
        if party in removed_set:
            status_rows.append({"Sold Party": party, "Generation_Status": "EXCLUDED", "Generation_Reason": "Removed party control"})
            continue
        behavior = pd.Series(behavior_map.get(party, {"Primary_Trigger": "HISTORICAL_PATTERN", "Behavior_Type": "SPECIFICATION-REPEAT"}))
        modes = p.determine_candidate_modes(party, row, behavior, mode_summary, overrides, settings)
        force_mode = "AUTO"
        if not overrides.empty and "Sold Party" in overrides.columns:
            ov = overrides[overrides["Sold Party"].fillna("").astype(str).str.strip().eq(party)]
            if not ov.empty:
                force_mode = p.clean_text(ov.iloc[-1].get("Force_Deal_Mode", "AUTO")).upper() or "AUTO"
        party_actions = action_index.get(party, current_actions.iloc[0:0] if not current_actions.empty else pd.DataFrame())
        current_action_stones = int(party_actions.get("Serial_Key", pd.Series(dtype=str)).nunique()) if not party_actions.empty else 0
        reserved_pairs = set(pair_owners) if allocation == "UNIQUE" and pair_policy == "RESERVE_PAIR_KEY" else None
        results = [
            p.select_for_mode(
                party, deal_mode, recipes, stock, current_actions, behavior, settings, exclusions,
                stock_index, party_actions,
                global_used if allocation == "UNIQUE" else None,
                reserved_pairs,
            )
            for deal_mode in modes
        ]
        chosen = p.choose_mode_result(results, current_actions, party, settings)
        if chosen is None:
            status_rows.append({
                "Sold Party": party, "Generation_Status": "REVIEW", "Generation_Reason": "No due mode/eligible recipe",
                "Interval_Status": row.get("Interval_Status"), "Behavior_Type": behavior.get("Behavior_Type"),
                "Mode_Selection_Override": force_mode if force_mode in p.MODE_ORDER else "NO",
                "Current_Action_Stone_Count": current_action_stones
            })
            continue

        chosen_selected_keys = set(chosen.selected.get("Serial_Key", pd.Series(dtype=str)).astype(str)) if not chosen.selected.empty else set()
        for result in results:
            if not result.line_status.empty:
                line_status = result.line_status.copy()
                line_status["Chosen_Mode"] = "YES" if result.mode == chosen.mode else "NO"
                line_parts.append(line_status)
            if not result.candidates.empty:
                candidates = result.candidates.copy()
                candidates["Chosen_Mode"] = "YES" if result.mode == chosen.mode else "NO"
                candidate_total += len(candidates)
                candidate_parts.append(candidates)
                preview_parts.append(candidates.sort_values(["Final_Selection_Score", "Action_Relevance_Score"], ascending=False).head(20))
                if result.mode == chosen.mode and "Direct_Action_Flag" in candidates.columns:
                    direct = candidates[candidates["Direct_Action_Flag"].eq("YES")].copy()
                    if not direct.empty:
                        direct["Selected_In_Final_Basket"] = direct["Serial_Key"].astype(str).isin(chosen_selected_keys).map({True: "YES", False: "NO"})
                        direct = direct[direct["Selected_In_Final_Basket"].eq("NO")]
                        if not direct.empty:
                            direct["Omission_Reason"] = direct.apply(
                                lambda r: "RECIPE LINE TARGET ZERO" if int(float(r.get("Target_Quantity", 0) or 0)) <= 0
                                else "NOT SELECTED AFTER DIRECT-FIRST/MATCH/SCORE ORDER OR GLOBAL RESERVATION",
                                axis=1,
                            )
                            omission_parts.append(direct)

        accepted = bool(chosen.accepted)
        pair_added_count = int(chosen.selected.get("Pair_Added_Flag", pd.Series("NO", index=chosen.selected.index)).astype(str).str.upper().eq("YES").sum()) if not chosen.selected.empty else 0
        normal_selected_count = len(chosen.selected) - pair_added_count
        status = "GENERATED" if accepted else "REVIEW - STOCK/COVERAGE/QUALITY/PAIR CONTROL"
        chosen_mode_rows = mode_summary[(mode_summary["Sold Party"].eq(party)) & (mode_summary["Deal_Mode"].eq(chosen.mode))]
        chosen_mode_row = chosen_mode_rows.iloc[-1] if not chosen_mode_rows.empty else pd.Series(dtype=object)
        action_info = p.classify_action_support(chosen.mode, current_action_stones, settings, actions_staleness_days)
        action_required = int(action_info["Action_Support_Threshold"])
        action_passed = action_info["Action_Support_Passed"] == "YES"
        status_rows.append({
            "Sold Party": party, "Sales_Person": row.get("Sales_Person"), "Interval_Status": row.get("Interval_Status"),
            "Behavior_Type": behavior.get("Behavior_Type"), "Primary_Trigger": behavior.get("Primary_Trigger"),
            "Chosen_Deal_Mode": chosen.mode, "Target_Stone_Count": chosen.target, "Selected_Stone_Count": len(chosen.selected),
            "Normal_Selected_Stone_Count": normal_selected_count, "Pair_Added_Stone_Count": pair_added_count,
            "Stock_Coverage": chosen.coverage, "Core_Stock_Coverage": chosen.core_coverage,
            "Chosen_Mode_Interval_Status": chosen_mode_row.get("Mode_Interval_Status", ""),
            "Chosen_Mode_Interval_Source": chosen_mode_row.get("Mode_Interval_Source", ""),
            "Chosen_Mode_Interval_Progress": chosen_mode_row.get("Mode_Interval_Progress", ""),
            "Mode_Selection_Override": force_mode if force_mode in p.MODE_ORDER else "NO",
            "Current_Action_Stone_Count": current_action_stones,
            "Mode_Action_Gate_Required": action_required,
            "Mode_Action_Gate_Passed": "YES" if p.action_policy(settings) == "BOOST_ONLY" else ("YES" if action_passed else "NO"),
            **action_info,
            "Generation_Status": status, "Generation_Reason": chosen.reason,
            "Global_Allocation_Mode": allocation, "Pair_Allocation_Policy": pair_policy,
        })
        if not chosen.selected.empty:
            selected = chosen.selected.copy()
            selected["Sales_Person"] = row.get("Sales_Person")
            selected["Interval_Status"] = row.get("Interval_Status")
            selected["Behavior_Type"] = behavior.get("Behavior_Type")
            selected["Chosen_Mode_Interval_Status"] = chosen_mode_row.get("Mode_Interval_Status", "")
            selected["Chosen_Mode_Interval_Source"] = chosen_mode_row.get("Mode_Interval_Source", "")
            selected["Mode_Selection_Override"] = force_mode if force_mode in p.MODE_ORDER else "NO"
            selected["Historical_Target_Quantity"] = chosen.target
            for key, value in action_info.items():
                selected[key] = value
            if accepted:
                if allocation == "UNIQUE":
                    global_used.update(selected["Serial_Key"].dropna().astype(str))
                    if pair_policy == "RESERVE_PAIR_KEY" and "Pair_Key" in selected.columns:
                        for pair_key in selected["Pair_Key"].fillna("").astype(str):
                            if pair_key:
                                pair_owners.setdefault(pair_key, party)
                selected_parts.append(selected)
            else:
                partial_parts.append(selected)

    # Candidate frames can have mode-specific or pair-completion audit columns.
    # Consolidate to the union schema, then serialize exactly once.
    candidate_df = write_candidate_batch(candidate_path, candidate_parts)
    if len(candidate_df) != candidate_total:
        raise RuntimeError(
            f"Candidate batch row-count mismatch: expected {candidate_total}, wrote {len(candidate_df)}"
        )

    status_df = pd.DataFrame(status_rows)
    selected_df = _concat_frames_no_future_warning(selected_parts)
    partial_df = _concat_frames_no_future_warning(partial_parts)
    line_df = _concat_frames_no_future_warning(line_parts)
    preview_df = _concat_frames_no_future_warning(preview_parts)
    omissions_df = _concat_frames_no_future_warning(omission_parts)
    if not omissions_df.empty:
        omissions_df = omissions_df.sort_values(["Sold Party", "Recipe_Line_ID", "Final_Selection_Score"], ascending=[True, True, False])
        omissions_df = omissions_df.drop_duplicates(["Sold Party", "Deal_Mode", "Serial_Key"], keep="first")

    for name, obj in [
        ("status", status_df), ("selected", selected_df), ("review_partial", partial_df),
        ("line_status", line_df), ("candidate_preview", preview_df),
        ("direct_action_omissions", omissions_df), ("candidate_total", candidate_total),
    ]:
        dump(out, f"{name}_batch_{args.chunk:03d}", obj)
    dump(out, f"allocation_state_{args.chunk:03d}", {"used_serials": global_used, "pair_owners": pair_owners})

    print(
        f"batch={args.chunk} range={start}:{end} generated={int(status_df.get('Generation_Status', pd.Series(dtype=str)).eq('GENERATED').sum())} "
        f"selected={len(selected_df)} reserved_serials={len(global_used)} reserved_pairs={len(pair_owners)} "
        f"candidates={candidate_total} seconds={time.time()-started:.1f}", flush=True
    )
    os._exit(0)


if __name__ == "__main__":
    main()
