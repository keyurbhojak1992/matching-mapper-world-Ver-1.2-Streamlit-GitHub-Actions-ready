# Validation report — version 1.4.3

## Completed checks

- Python compilation completed for the full repository.
- Automated test suite: **21 passed** in an isolated environment using the pinned project dependencies.
- The strict single-sheet regression test deliberately sets `MAX_ROWS_PER_SHEET = 1`; the final global master still contains exactly one worksheet named `Master_Data`.
- Production-scale validation used the user's 28 available batch master workbooks and an intermediate limit of 13,000 rows:
  - source master workbooks: 28;
  - final worksheets: 1;
  - worksheet: `Master_Data`;
  - data rows: 71,801;
  - total rows including header: 71,802;
  - columns: 58;
  - formula hyperlinks: 358,999;
  - native hyperlink objects: 0;
  - worksheet dimension: `A1:BF71802`;
  - XLSX archive creation: passed.
- The final assembler ignores the intermediate `MAX_ROWS_PER_SHEET` value and uses Excel's hard maximum only for the user-facing consolidated workbook.
- If a future result exceeds 1,048,575 data rows plus one header row, the worker raises a clear error instead of silently creating `Master_Data_Part_N` tabs.
- The final manifest records `assembler_version: 1.4.3`, `sheet_policy: strict_single_sheet`, and `sheet_name: Master_Data`.
- All eight protected files under `original_master_code/` pass `SOURCE_CHECKSUMS.sha256` unchanged.

## Deployment note

A completed job created by an older worker is immutable. After deploying version 1.4.3, submit a **new** GitHub Actions job; recovering or redownloading an older release will return the older multi-sheet workbook.
