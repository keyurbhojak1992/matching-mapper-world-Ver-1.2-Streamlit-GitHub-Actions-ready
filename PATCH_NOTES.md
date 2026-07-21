# Version 1.4.3 strict single-sheet hotfix

- Forces the final `MASTER_CONSOLIDATED_REPORT.xlsx` into exactly one `Master_Data` worksheet whenever the data fits Excel's hard row limit.
- Prevents the intermediate `MAX_ROWS_PER_SHEET` setting from creating `Master_Data_Part_N` tabs in the final output.
- Adds an explicit assembler version and sheet policy to the manifest and final-output readme.
- Raises a clear error rather than silently splitting when more than 1,048,575 data rows are requested.

# Version 1.4.1 hotfix

- Prevents `ImportError: cannot import name LEGACY_RESULT_ASSET_NAME` when a repository contains a mixture of v1.2 and v1.4 files.
- Keeps result recovery compatible with both `suggested_reports_all_batches.zip` and `suggested_reports_final_combined.zip`.
- The full repository should still be replaced together so the worker, workflow, adapter, and UI stay on the same version.

# Matching Mapper World — final-output patch (v1.4)

## Fixed

- New GitHub Actions full runs no longer return nested `batch_XX/suggested_reports.zip` packages as the user-facing result.
- Batch-level `MASTER_CONSOLIDATED_REPORT.xlsx` workbooks are assembled into one global master workbook.
- When **Generate salesperson workbooks** is enabled, every salesperson workbook is placed together at the root of the final ZIP.
- When **Create separated party ZIPs** is enabled, every salesperson-specific separated ZIP is also placed at the root of the final ZIP.
- Batch logs and runtime snapshots are retained separately under `_batch_diagnostics/`.
- A completed job now retains `current_batch` and `total_batches`, so the dashboard shows `30/30` rather than `Waiting` after a 30-batch run.

## Final output layout

```text
suggested_reports_final_combined.zip
├── MASTER_CONSOLIDATED_REPORT.xlsx       # when merged master is enabled
├── Salesperson A.xlsx                    # when salesperson workbooks are enabled
├── Salesperson B.xlsx
├── Salesperson A_Separated_Parties.zip   # when separated output is enabled
├── ...
├── _batch_diagnostics/
├── batch_manifest.json
├── processing_log.txt
└── FINAL_OUTPUT_README.txt
```

The merged master uses low-memory OOXML streaming. Columns are aligned by exact header name, original row-fill rules are recreated, and decimal formats and widths are retained. Version 1.4.2 writes links as `HYPERLINK()` formulas and keeps the merged master in one worksheet named `Master_Data` when the row count fits in one Excel sheet.

## Preserved

- All eight files under `original_master_code/` are byte-for-byte unchanged and match `SOURCE_CHECKSUMS.sha256`.
- Every batch receives the complete selected configuration.
- Only `TEST_MODE=True` and `TEST_SALESPEOPLE=current_batch` are changed per batch.
- Local/small-test processing remains available and is limited to a small selected scope to protect Streamlit Community Cloud.

## Important for individual files

The output toggle **Generate salesperson workbooks** must be on before job submission. A completed job cannot recreate those workbooks later when the submitted runtime had `GENERATE_INDIVIDUAL_REPORTS=false` because they were never produced by the preserved master engine.
