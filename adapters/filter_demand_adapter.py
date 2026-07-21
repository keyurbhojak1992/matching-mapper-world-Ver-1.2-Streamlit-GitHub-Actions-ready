from __future__ import annotations

from datetime import datetime

from .common import (
    MASTER_CODE_DIR,
    apply_runtime_overrides,
    captured_output,
    load_fresh_module,
    temp_workspace,
    write_bytes,
)

MASTER_PATH = MASTER_CODE_DIR / "Filter_Demand_Stock_Process.py"


def run_filter_demand(
    *,
    stock_bytes: bytes,
    filter_bytes: bytes,
    runtime_overrides: dict[str, object] | None = None,
) -> tuple[bytes, str, str]:
    with temp_workspace("matching_mapper_filter_") as work:
        stock_path = write_bytes(work / "stock Actual.xlsx", stock_bytes)
        filter_path = write_bytes(work / "filters.xlsx", filter_bytes)
        output_name = f"Demand_Filter_Matches_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.zip"
        output_path = work / output_name

        module = load_fresh_module(MASTER_PATH, "filter_demand_master")
        apply_runtime_overrides(module, runtime_overrides)

        # Rebuild constants that the untouched source computes at import time.
        if hasattr(module, "ENABLE_FLO_MAPPING"):
            module.FLUORESCENCE_ORDER = (
                ["NON", "FNT", "MED", "STG", "VST"]
                if module.ENABLE_FLO_MAPPING
                else ["NON", "VSL", "SL", "FNT", "MED", "STG", "VST"]
            )
            if hasattr(module, "ORDER_MAPPING"):
                module.ORDER_MAPPING["FLO"] = module.FLUORESCENCE_ORDER

        with captured_output() as logs:
            module.generate_salesperson_reports(
                str(stock_path),
                str(output_path),
                str(filter_path),
                int(getattr(module, "TOLERANCE_LEVEL", 1)),
            )

        if not output_path.exists():
            raise RuntimeError(
                "No output ZIP was created. Confirm that the filter workbook contains "
                "a valid 'Change Format' sheet and salesperson rows.\n\n" + logs.getvalue()
            )

        return output_path.read_bytes(), output_name, logs.getvalue()
