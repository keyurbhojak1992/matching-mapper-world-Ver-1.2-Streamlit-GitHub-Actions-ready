from __future__ import annotations

import io
import re
import tempfile
import zipfile
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path

from adapters.common import MASTER_CODE_DIR


MASTER_SOURCE = MASTER_CODE_DIR / "Auto_Pricing.py"
ACTION_NAMES = ("View", "Hold", "Offer", "Wishlist", "Cart", "Media")


def _runtime_namespace(output_main: Path, output_verify: Path) -> dict[str, object]:
    """Load the preserved source with only its two desktop output paths redirected."""
    source = MASTER_SOURCE.read_text(encoding="utf-8")
    source = source.split("# --- Execution ---", 1)[0]

    replacements = [
        (
            r"(?m)^\s*df_upgraded\.to_excel\([^\n]+index=False\)\s*$",
            f"    df_upgraded.to_excel({str(output_main)!r}, index=False)",
        ),
        (
            r"(?m)^\s*df_verify\.to_excel\([^\n]+index=False\)\s*$",
            f"    df_verify.to_excel({str(output_verify)!r}, index=False)",
        ),
    ]
    for pattern, replacement in replacements:
        source, count = re.subn(pattern, replacement, source, count=1)
        if count != 1:
            raise RuntimeError("Auto Pricing source changed: output redirect point was not found exactly once.")

    namespace: dict[str, object] = {
        "__name__": "matching_mapper_auto_pricing_runtime",
        "__file__": str(MASTER_SOURCE),
    }
    exec(compile(source, str(MASTER_SOURCE), "exec"), namespace, namespace)
    if not callable(namespace.get("process_stock_upgrade")):
        raise RuntimeError("Auto Pricing source does not expose process_stock_upgrade().")
    return namespace


def run_auto_pricing(
    *,
    stock_bytes: bytes,
    sales_180_bytes: bytes,
    sales_120_bytes: bytes,
    action_files: dict[str, bytes | None],
) -> tuple[bytes, str, str]:
    """Run the preserved Auto Pricing source and return its two XLSX outputs in one ZIP."""
    with tempfile.TemporaryDirectory(prefix="matching_mapper_auto_pricing_") as tmp:
        work = Path(tmp)
        stock_path = work / "Stock file.xlsx"
        sales_180_path = work / "sales_180_file.xlsx"
        sales_120_path = work / "sales_120_file.xlsx"
        stock_path.write_bytes(stock_bytes)
        sales_180_path.write_bytes(sales_180_bytes)
        sales_120_path.write_bytes(sales_120_bytes)

        action_paths: dict[str, str] = {}
        for action_name in ACTION_NAMES:
            payload = action_files.get(action_name)
            path = work / f"{action_name}.xlsx"
            if payload is not None:
                path.write_bytes(payload)
            # Keep all six action keys, as in the supplied execution block. A missing
            # upload intentionally points to a non-existent file; the preserved source
            # catches that read error after already creating the action count columns.
            action_paths[action_name] = str(path)

        output_main = work / "Auto pricing revision.xlsx"
        output_verify = work / "Cross_Verify_Sales_Actions.xlsx"
        namespace = _runtime_namespace(output_main, output_verify)

        captured = io.StringIO()
        with redirect_stdout(captured), redirect_stderr(captured):
            namespace["process_stock_upgrade"](
                str(stock_path),
                str(sales_180_path),
                str(sales_120_path),
                action_paths,
            )

        if not output_main.is_file() or not output_verify.is_file():
            raise RuntimeError(
                "Auto Pricing completed without creating both expected workbooks.\n"
                + captured.getvalue()[-6000:]
            )

        output_zip = work / "Auto_Pricing_Output.zip"
        with zipfile.ZipFile(output_zip, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.write(output_main, arcname=output_main.name)
            zf.write(output_verify, arcname=output_verify.name)

        return output_zip.read_bytes(), output_zip.name, captured.getvalue()
