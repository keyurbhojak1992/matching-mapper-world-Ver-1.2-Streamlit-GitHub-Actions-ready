from __future__ import annotations

import gc
import shutil
import sys
from datetime import datetime
from pathlib import Path

from .common import (
    MASTER_CODE_DIR,
    captured_output,
    load_fresh_module,
    resolve_repository_input,
    temp_workspace,
    write_bytes,
)

MASTER_PATH = MASTER_CODE_DIR / "BID_Drop_Checking_Process.py"


def _materialize_rule_file(
    *,
    work: Path,
    uploaded_bytes: bytes | None,
    default_key: str,
    destination_name: str,
) -> Path:
    destination = work / destination_name
    if uploaded_bytes is not None:
        return write_bytes(destination, uploaded_bytes)
    source = resolve_repository_input(None, default_key, required=True)
    assert source is not None
    shutil.copy2(source, destination)
    return destination


def run_bid_drop_checking(
    *,
    suggested_bytes: bytes,
    bids_bytes: bytes,
    sales_bytes: bytes,
    stock_bytes: bytes,
    filter_bytes: bytes | None,
    no_need_bytes: bytes | None,
) -> tuple[bytes, str, str]:
    """Run the preserved BID drop diagnostic with upload/default file handling."""
    with temp_workspace("matching_mapper_bid_drop_") as work:
        suggested_path = write_bytes(work / "MASTER_CONSOLIDATED_REPORT.xlsx", suggested_bytes)
        bids_path = write_bytes(work / "bids.xlsx", bids_bytes)
        sales_path = write_bytes(work / "Sales and Actions.xlsx", sales_bytes)
        stock_path = write_bytes(work / "Stock Actual.xlsx", stock_bytes)
        filters_path = _materialize_rule_file(
            work=work,
            uploaded_bytes=filter_bytes,
            default_key="filters",
            destination_name="filters - EBID.xlsx",
        )
        no_need_path = _materialize_rule_file(
            work=work,
            uploaded_bytes=no_need_bytes,
            default_key="no_need",
            destination_name="no_need.xlsx",
        )

        output_name = f"Deep_Gap_Analysis_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
        output_path = work / output_name
        module = None
        try:
            module = load_fresh_module(MASTER_PATH, "bid_drop_master")
            with captured_output() as logs:
                module.run_bid_drop_check(
                    suggested_path=suggested_path,
                    bids_path=bids_path,
                    sales_path=sales_path,
                    stock_path=stock_path,
                    filters_path=filters_path,
                    no_need_path=no_need_path,
                    output_path=output_path,
                )
            if not output_path.is_file():
                raise RuntimeError("BID Drop Checking completed without creating the analysis workbook.")
            return output_path.read_bytes(), output_name, logs.getvalue()
        finally:
            if module is not None:
                sys.modules.pop(module.__name__, None)
                del module
            gc.collect()
