# Hotfix 1.4.2 — single-sheet merged master

This patch changes the final merged master output so `MASTER_CONSOLIDATED_REPORT.xlsx` is created as a single workbook with a single worksheet named `Master_Data` whenever the total row count fits inside Excel's worksheet row limit.

Why this changed:
- Version 1.4 split the master into `Master_Data_Part_N` sheets to stay below Excel's native hyperlink-per-worksheet limit.
- Version 1.4.2 keeps all rows on one sheet by writing report links as `HYPERLINK()` formulas instead of native hyperlink relationships.
- The protected source scripts in `original_master_code/` remain unchanged.

Expected final ZIP layout:

```text
suggested_reports_final_combined.zip
├── MASTER_CONSOLIDATED_REPORT.xlsx   # one worksheet: Master_Data
├── salesperson workbooks, when enabled
├── separated-party ZIPs, when enabled
├── _batch_diagnostics/
├── batch_manifest.json
└── processing_log.txt
```

For the supplied production master, the regenerated single-sheet workbook contains:

```text
Worksheet: Master_Data
Data rows: 71,801
Columns: 58
Formula hyperlinks: 358,999
Native hyperlink objects: 0
```
