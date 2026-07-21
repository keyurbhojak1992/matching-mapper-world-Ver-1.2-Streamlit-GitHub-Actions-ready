# Hotfix 1.4.3 — strict single-sheet final master

This patch closes the remaining path that could still create `Master_Data_Part_N` tabs.

- `MAX_ROWS_PER_SHEET` now applies only to the preserved engine's intermediate batch workbooks.
- The final downloaded `MASTER_CONSOLIDATED_REPORT.xlsx` ignores that intermediate limit.
- The final master is created with exactly one worksheet named `Master_Data` whenever the full result has at most 1,048,575 data rows plus one header row.
- Report links use `HYPERLINK()` formulas, so Excel's native hyperlink-object limit cannot force sheet splitting.
- If a future result exceeds Excel's hard one-worksheet row limit, the worker now reports a clear error instead of silently creating multiple tabs.
- `batch_manifest.json` records `assembler_version: 1.4.3` and `sheet_policy: strict_single_sheet`.
- `FINAL_OUTPUT_README.txt` also records the assembler version and sheet policy so the downloaded result can be verified.

The files in `original_master_code/` remain unchanged.
