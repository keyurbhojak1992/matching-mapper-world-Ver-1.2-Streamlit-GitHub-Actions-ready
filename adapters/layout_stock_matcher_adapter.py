from __future__ import annotations

import sys
import tempfile
from pathlib import Path

from adapters.common import MASTER_CODE_DIR, captured_output, load_fresh_module


MASTER_SOURCE = MASTER_CODE_DIR / "Layout_Stock_Matcher.py"


def run_layout_stock_matcher(
    *,
    layout_bytes: bytes,
    sales_bytes: bytes,
    actions_bytes: bytes,
    tolerance: int = 1,
    min_layout_pct: float = 0.10,
    min_absolute_stones: int = 5,
    min_action_weight: int = 45,
    include_match_columns: bool = True,
    create_empty_salesperson_files: bool = True,
) -> tuple[bytes, str, str]:
    """Invoke the preserved Layout Stock Matcher build_outputs() function."""
    module = load_fresh_module(MASTER_SOURCE, "layout_stock_matcher")
    try:
        with tempfile.TemporaryDirectory(prefix="matching_mapper_layout_") as tmp:
            work = Path(tmp)
            layout_path = work / "Layout Stock.xlsx"
            sales_path = work / "Sales.xlsx"
            actions_path = work / "Actions.xlsx"
            output_zip = work / module.DEFAULT_OUTPUT_ZIP
            layout_path.write_bytes(layout_bytes)
            sales_path.write_bytes(sales_bytes)
            actions_path.write_bytes(actions_bytes)

            config = module.Config(
                tolerance=int(tolerance),
                min_layout_pct=float(min_layout_pct),
                min_absolute_stones=int(min_absolute_stones),
                min_action_weight=int(min_action_weight),
                include_match_columns=bool(include_match_columns),
                create_empty_salesperson_files=bool(create_empty_salesperson_files),
            )

            with captured_output() as log:
                stats = module.build_outputs(
                    layout_path=layout_path,
                    sales_path=sales_path,
                    actions_path=actions_path,
                    output_zip=output_zip,
                    config=config,
                )
                print("Layout Stock Matcher stats:", stats)

            if not output_zip.is_file():
                raise RuntimeError("Layout Stock Matcher completed without creating its output ZIP.")
            return output_zip.read_bytes(), output_zip.name, log.getvalue()
    finally:
        sys.modules.pop(module.__name__, None)
