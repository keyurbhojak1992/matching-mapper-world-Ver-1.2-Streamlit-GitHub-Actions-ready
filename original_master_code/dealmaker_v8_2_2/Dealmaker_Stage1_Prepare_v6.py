from __future__ import annotations

import argparse
import gc
import importlib.util
import json
import pickle
import sys
import time
from pathlib import Path

import pandas as pd


def dump(root: Path, name: str, obj) -> None:
    with open(root / f"{name}.pkl", "wb") as fh:
        pickle.dump(obj, fh, pickle.HIGHEST_PROTOCOL)


def mark(text: str, started: float) -> None:
    print(f"stage1a [{time.time()-started:7.1f}s] {text}", flush=True)


def main() -> None:
    ap = argparse.ArgumentParser(description="Build production deal history, intervals, behavior, and current-action checkpoints.")
    ap.add_argument("--sales", required=True)
    ap.add_argument("--actions", required=True)
    ap.add_argument("--control", required=True)
    ap.add_argument("--pipeline", required=True)
    ap.add_argument("--interval-engine", required=True)
    ap.add_argument("--work-dir", required=True)
    ap.add_argument("--sales-sheet", default="")
    ap.add_argument("--actions-sheet", default="")
    ap.add_argument("--sales-header-row", type=int, default=None)
    ap.add_argument("--actions-header-row", type=int, default=None)
    args = ap.parse_args()

    started = time.time()
    pipeline_path = Path(args.pipeline).resolve()
    sys.path.insert(0, str(pipeline_path.parent))
    spec = importlib.util.spec_from_file_location("dealmaker_pipeline_v6_stage1a", pipeline_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Could not load pipeline: {pipeline_path}")
    p = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = p
    spec.loader.exec_module(p)
    interval = p.load_module(Path(args.interval_engine), "dealmaker_interval_v10_stage1a")

    work = Path(args.work_dir)
    work.mkdir(parents=True, exist_ok=True)
    control = Path(args.control)
    settings, overrides, removed, exclusions = p.load_control(control)
    quality_config = p.configure_quality(control)
    if hasattr(interval, "configure_quality"):
        interval.configure_quality(control)
    if any(i.get("Severity") == "ERROR" for i in quality_config.validation_issues):
        raise ValueError(f"Invalid quality control: {quality_config.validation_issues}")

    quality: list[dict] = []
    aliases = p.load_party_aliases(control)
    mark("reading Sales", started)
    sales = interval.prepare_dataset(
        Path(args.sales), "Sales", quality,
        args.sales_sheet or None, args.sales_header_row,
        p.load_manual_mapping(control, "SALES"),
    )
    mark("reading Actions", started)
    actions = interval.prepare_dataset(
        Path(args.actions), "Actions", quality,
        args.actions_sheet or None, args.actions_header_row,
        p.load_manual_mapping(control, "ACTIONS"),
    )
    sales = p.apply_party_aliases(sales, aliases, "Sales", quality)
    actions = p.apply_party_aliases(actions, aliases, "Actions", quality)
    sales = interval.dedupe_business_rows(sales, "Sales", quality)
    actions = interval.dedupe_business_rows(actions, "Actions", quality)
    actions = p.canonicalize_action_types(actions, settings)

    as_of = interval.determine_as_of(settings, sales, actions)
    sale_types = p.parse_csv_set(settings.get("SALE_TYPE_VALUES", "SALE,SALES"))
    bid_types = p.parse_csv_set(settings.get("BID_TYPE_VALUES", "BID"))
    sales_only = sales[sales["Type"].isin(sale_types)].copy()
    bids = sales[sales["Type"].isin(bid_types)].copy()
    if sales_only.empty:
        raise ValueError(f"No completed SALE rows found. SALE_TYPE_VALUES={sorted(sale_types)}")

    mark("building deal buckets", started)
    deals, _ = interval.build_deal_buckets(
        sales_only,
        int(float(settings.get("MIN_DEAL_STONES", 6))),
        int(float(settings.get("MAX_DEAL_DAYS", 2))),
    )
    party_summary, _ = interval.build_party_summary(
        sales, sales_only, bids, actions, deals, overrides, removed, settings, as_of
    )
    mode_summary = p.build_mode_summary(deals, party_summary, settings, as_of)

    sales_rows = len(sales)
    sale_rows = len(sales_only)
    bid_rows = len(bids)
    action_rows = len(actions)
    hold_action_rows = int(actions["Type"].eq("HOLD").sum())

    # Sales and BID totals are checkpointed as metadata; the wide frames are
    # released before action conversion to keep peak memory bounded.
    del sales, bids
    gc.collect()

    mark("building party action conversion", started)
    action_conversion = p.build_action_conversion(
        actions, sales_only, deals,
        int(float(settings.get("ACTION_CONVERSION_WINDOW_DAYS", 30))),
    )
    mark("building behavior profiles", started)
    behavior = p.build_behavior_profiles(party_summary, mode_summary, action_conversion, settings)
    mark("joining deal membership", started)
    detail = p.add_deal_membership(sales_only, deals)
    mark("enriching deal pattern fields", started)
    detail = p.enrich_pattern_fields(detail)
    mark("auditing repeated deal serials", started)
    repeated_deal_serial_audit = p.build_repeated_deal_serial_audit(detail)
    profile_detail, repeated_deal_serial_rows = p.dedupe_deal_detail_for_profiles(detail)
    quality.append({
        "Dataset": "Deals",
        "Check": "Repeated serial rows within the same deal removed from profile calculations",
        "Value": int(repeated_deal_serial_rows),
        "Severity": "WARNING" if repeated_deal_serial_rows else "OK",
    })
    mark("building current actions", started)
    current_actions = p.build_current_actions(actions, deals, action_conversion, settings, as_of)
    mark("serializing history checkpoints", started)

    sold_serial_keys = set(sales_only["Serial_Key"].astype(str))
    latest_sale_date = sales_only["Date"].max()
    latest_action_date = actions["Date"].max() if not actions.empty else pd.NaT
    partial_metadata = {
        "pipeline_version": p.VERSION,
        "history_elapsed_seconds": round(time.time() - started, 2),
        "sales_rows": sales_rows,
        "sale_rows": sale_rows,
        "bid_rows": bid_rows,
        "action_rows": action_rows,
        "deal_buckets": len(deals),
        "parties": len(party_summary),
        "eligible_parties": int(party_summary["Auto_Include"].astype(str).str.upper().eq("YES").sum()),
        "repeated_deal_serial_rows_removed": int(repeated_deal_serial_rows),
        "repeated_deal_serial_groups": int(len(repeated_deal_serial_audit)),
        "hold_action_rows": hold_action_rows,
        "action_types": sorted(actions["Type"].dropna().astype(str).unique().tolist()),
        "latest_sale_date": str(latest_sale_date.date()) if pd.notna(latest_sale_date) else None,
        "latest_action_date": str(latest_action_date.date()) if pd.notna(latest_action_date) else None,
    }

    objects = {
        "settings": settings,
        "overrides": overrides,
        "removed": removed,
        "exclusions": exclusions,
        "aliases": aliases,
        "as_of": as_of,
        "deals": deals,
        "party_summary": party_summary,
        "mode_summary": mode_summary,
        "action_conversion": action_conversion,
        "behavior_profiles": behavior,
        "profile_detail": profile_detail,
        "repeated_deal_serial_audit": repeated_deal_serial_audit,
        "current_actions": current_actions,
        "sold_serial_keys": sold_serial_keys,
        "quality_log": quality,
        "quality_domain_config": quality_config.domain,
        "quality_alias_config": quality_config.aliases,
    }
    for name, obj in objects.items():
        dump(work, name, obj)
    (work / "stage1_partial_metadata.json").write_text(
        json.dumps(partial_metadata, indent=2, default=str), encoding="utf-8"
    )
    mark("history checkpoint complete", started)
    print(json.dumps(partial_metadata, indent=2), flush=True)


if __name__ == "__main__":
    main()
