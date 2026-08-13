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


def load(root: Path, name: str):
    with open(root / f"{name}.pkl", "rb") as fh:
        return pickle.load(fh)


def dump(root: Path, name: str, obj) -> None:
    with open(root / f"{name}.pkl", "wb") as fh:
        pickle.dump(obj, fh, pickle.HIGHEST_PROTOCOL)


def raw_quality_view(frame: pd.DataFrame) -> pd.DataFrame:
    fields = ["Shade", "Milky", "TOIN_NAME", "SOIN_NAME", "GIRDLE_OPEN_NAME", "PAV_OPEN_NAME", "TBINC_NAME", "SBINC_NAME"]
    out = pd.DataFrame(index=frame.index)
    for field in fields:
        raw = f"Raw_{field}"
        out[field] = frame[raw] if raw in frame.columns else frame.get(field, "")
    return out


def main() -> None:
    ap = argparse.ArgumentParser(description="Build combination profiles, recipes, and eligible production stock in a fresh process.")
    ap.add_argument("--stock", required=True)
    ap.add_argument("--control", required=True)
    ap.add_argument("--pipeline", required=True)
    ap.add_argument("--interval-engine", required=True)
    ap.add_argument("--work-dir", required=True)
    ap.add_argument("--stock-sheet", default="")
    ap.add_argument("--stock-header-row", type=int, default=None)
    args = ap.parse_args()

    started = time.time()
    pipeline_path = Path(args.pipeline).resolve()
    sys.path.insert(0, str(pipeline_path.parent))
    spec = importlib.util.spec_from_file_location("dealmaker_pipeline_v6_stage1b", pipeline_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Could not load pipeline: {pipeline_path}")
    p = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = p
    spec.loader.exec_module(p)
    interval = p.load_module(Path(args.interval_engine), "dealmaker_interval_v10_stage1b")

    work = Path(args.work_dir)
    control = Path(args.control)
    settings = load(work, "settings")
    quality = load(work, "quality_log")
    quality_config = p.configure_quality(control)
    if hasattr(interval, "configure_quality"):
        interval.configure_quality(control)

    deals = load(work, "deals")
    mode_summary = load(work, "mode_summary")
    profile_detail = load(work, "profile_detail")
    combinations = p.build_combination_profiles(profile_detail, deals, settings)
    recipes = p.build_recipe_lines(combinations, mode_summary, settings)
    include_due_soon = p.yes(settings.get("INCLUDE_DUE_SOON", "NO"))
    allowed_mode_statuses = {"DUE NOW", "OVERDUE"} | ({"DUE SOON"} if include_due_soon else set())
    mode_due_eligible_parties = int(mode_summary[mode_summary["Mode_Interval_Status"].isin(allowed_mode_statuses)]["Sold Party"].nunique())
    dump(work, "combination_profiles", combinations)
    dump(work, "recipes", recipes)
    del combinations, recipes, profile_detail, mode_summary, deals
    gc.collect()

    stock, duplicate_stock, unavailable_stock = p.prepare_stock(
        Path(args.stock), interval, settings, quality,
        args.stock_sheet or None, args.stock_header_row,
        p.load_manual_mapping(control, "STOCK"),
    )
    stock_audit = p.quality_domain.domain_audit(raw_quality_view(stock), "Stock", quality_config)
    sold_serial_keys = load(work, "sold_serial_keys")
    historical_sale_stock = stock[stock["Serial_Key"].astype(str).isin(sold_serial_keys)].copy()
    pair_sale_excluded = stock.iloc[0:0].copy()
    if p.yes(settings.get("EXCLUDE_COMPLETED_SALE_SERIALS_FROM_STOCK", "YES")) and not historical_sale_stock.empty:
        stock = stock[~stock["Serial_Key"].astype(str).isin(set(historical_sale_stock["Serial_Key"].astype(str)))].copy()
        stock, pair_sale_excluded = p.cascade_pair_exclusions(
            stock,
            historical_sale_stock,
            "PAIR REMOVED: RELATED MEMBER FOUND IN COMPLETED-SALE HISTORY",
        )
        if not pair_sale_excluded.empty:
            unavailable_stock = pd.concat([unavailable_stock, pair_sale_excluded], ignore_index=True, sort=False)
        stock = stock.reset_index(drop=True)
    quality.extend([
        {
            "Dataset": "Stock",
            "Check": "Completed-SALE history serials excluded from candidate stock",
            "Value": len(historical_sale_stock),
            "Severity": "WARNING" if len(historical_sale_stock) else "OK",
        },
        {
            "Dataset": "Stock",
            "Check": "Related pair members excluded with completed-sale serials",
            "Value": len(pair_sale_excluded),
            "Severity": "WARNING" if len(pair_sale_excluded) else "OK",
        },
    ])

    partial = json.loads((work / "stage1_partial_metadata.json").read_text(encoding="utf-8"))
    run_date = pd.Timestamp.now(tz="Asia/Kolkata").tz_localize(None).normalize()
    latest_sale_date = pd.to_datetime(partial.get("latest_sale_date"), errors="coerce")
    latest_action_date = pd.to_datetime(partial.get("latest_action_date"), errors="coerce")
    sales_staleness_days = int((run_date - latest_sale_date).days) if pd.notna(latest_sale_date) else None
    actions_staleness_days = int((run_date - latest_action_date).days) if pd.notna(latest_action_date) else None
    stock_status_present = bool(stock.get("Stock_Status", pd.Series(dtype=str)).fillna("").astype(str).str.strip().ne("").any())

    for name, obj in {
        "quality_log": quality,
        "stock": stock,
        "historical_sale_stock": historical_sale_stock,
        "duplicate_stock": duplicate_stock,
        "unavailable_stock": unavailable_stock,
        "quality_domain_audit_stock": stock_audit,
    }.items():
        dump(work, name, obj)

    metadata = dict(partial)
    metadata.update({
        "elapsed_seconds": round(time.time() - started, 2),
        "stock_rows": len(stock),
        "eligible_parties": mode_due_eligible_parties if p.yes(settings.get("MODE_DUE_REQUIRED_FOR_SELECTION", "YES")) else int(partial.get("eligible_parties", 0)),
        "mode_due_eligible_parties": mode_due_eligible_parties,
        "historical_stock_overlaps": len(historical_sale_stock),
        "stock_status_present": stock_status_present,
        "run_date": str(run_date.date()),
        "sales_staleness_days": sales_staleness_days,
        "actions_staleness_days": actions_staleness_days,
        "quality_domain_errors": sum(1 for i in quality_config.validation_issues if i.get("Severity") == "ERROR"),
        "quality_domain_warnings": sum(1 for i in quality_config.validation_issues if i.get("Severity") == "WARNING"),
    })
    (work / "stage1_metadata.json").write_text(json.dumps(metadata, indent=2, default=str), encoding="utf-8")
    print(json.dumps(metadata, indent=2, default=str), flush=True)


if __name__ == "__main__":
    main()
