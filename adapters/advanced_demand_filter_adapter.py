from __future__ import annotations

from datetime import datetime

from processors.advanced_demand_batch import generate_salesperson_reports_batched

from .common import (
    MASTER_CODE_DIR,
    apply_runtime_overrides,
    captured_output,
    load_fresh_module,
    temp_workspace,
    write_bytes,
)

MASTER_PATH = MASTER_CODE_DIR / "Advance_Demand_Filter_Process.py"


def run_advanced_demand_filter(
    *,
    stock_bytes: bytes,
    filter_bytes: bytes,
    clean_bytes: bytes | None = None,
    runtime_overrides: dict[str, object] | None = None,
    min_stones_per_party: int = 0,
    output_layout: str = "salesperson_workbooks",
) -> tuple[bytes, str, str]:
    """Run the isolated Advanced Demand Filter workflow.

    Stock Actual and Filters are required uploads. No Need is optional; if it is
    not supplied, the clean/exclusion stage is skipped. No repository defaults
    are used by this workflow.
    """
    with temp_workspace("matching_mapper_advanced_demand_") as work:
        stock_path = write_bytes(work / "stock Actual.xlsx", stock_bytes)
        filter_path = write_bytes(work / "filters.xlsx", filter_bytes)
        clean_path = write_bytes(work / "no_need.xlsx", clean_bytes) if clean_bytes else None
        output_name = f"Advanced_Demand_Filter_Matches_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.zip"
        output_path = work / output_name

        module = load_fresh_module(MASTER_PATH, "advanced_demand_filter_master")
        apply_runtime_overrides(module, runtime_overrides)

        # The source builds this order at import time, so rebuild it after UI overrides.
        module.FLUORESCENCE_ORDER = (
            ["NON", "FNT", "MED", "STG", "VST"]
            if bool(getattr(module, "ENABLE_FLO_MAPPING", False))
            else ["NON", "VSL", "SL", "FNT", "MED", "STG", "VST"]
        )
        if hasattr(module, "ORDER_MAPPING"):
            module.ORDER_MAPPING["FLO"] = module.FLUORESCENCE_ORDER

        with captured_output() as logs:
            generate_salesperson_reports_batched(
                module,
                str(stock_path),
                str(output_path),
                str(filter_path),
                int(getattr(module, "TOLERANCE_LEVEL", 1)),
                str(clean_path) if clean_path is not None else None,
                min_stones_per_party=min_stones_per_party,
                output_layout=output_layout,
            )

        if not output_path.exists():
            raise RuntimeError(
                "No output ZIP was created. Confirm that Stock Actual has Sheet1 and the Filters "
                "workbook has a valid Change Format sheet with Sold Party and Sales_Person.\n\n"
                + logs.getvalue()
            )

        return output_path.read_bytes(), output_name, logs.getvalue()
