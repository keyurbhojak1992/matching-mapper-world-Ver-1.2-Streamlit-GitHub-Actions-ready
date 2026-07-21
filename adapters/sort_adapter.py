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

MASTER_PATH = MASTER_CODE_DIR / "Sorting_Suggested_Output_Zip_Process.py"


def run_sorting_process(
    *,
    input_zip_bytes: bytes,
    sales_file_bytes: bytes | None,
    runtime_overrides: dict[str, object] | None = None,
) -> tuple[bytes, str, str]:
    with temp_workspace("matching_mapper_sort_") as work:
        input_path = write_bytes(work / "suggested_input.zip", input_zip_bytes)
        sales_path = work / "sales_history.xlsx"
        if sales_file_bytes:
            write_bytes(sales_path, sales_file_bytes)

        output_name = f"Sorted_Suggested_Output_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.zip"
        output_path = work / output_name
        work_output_folder = work / "sort_work"

        module = load_fresh_module(MASTER_PATH, "sorting_master")
        apply_runtime_overrides(module, runtime_overrides)
        module.INPUT_ZIP_OR_FOLDER = str(input_path)
        module.SALES_FILE = str(sales_path)
        module.OUTPUT_FOLDER = str(work_output_folder)
        module.OUTPUT_ZIP_NAME = str(output_path)

        with captured_output() as logs:
            module.main()

        if not output_path.exists():
            raise RuntimeError("Sorting completed without an output ZIP.\n\n" + logs.getvalue())

        return output_path.read_bytes(), output_name, logs.getvalue()
