# Runtime Configuration Coverage

The original master scripts remain unchanged in `original_master_code/`. The Streamlit UI passes selected values to those modules only at runtime.

## Suggested Stock Mapper

Every operational variable from the original USER INPUT / OUTPUT CONTROL blocks is represented in the app:

### Matching

- `TOP_N_SUPERTAGS`
- `TOLERANCE_LEVEL`
- `ENABLE_FLO_MAPPING`
- All `INCLUDE_*` dimension switches
- All `ENABLE_*_TOLERANCE` switches

### Supertags and display

- `SHOW_BASIC_SUPERTAG`
- `SHOW_ADVANCED_SUPERTAG`
- `SHOW_REF_STONE_ID`
- `SHOW_CONCAT_STATUS`
- `ENABLE_TYPE_COLUMNS`
- `STATIC_TYPE_SEQUENCE`
- `ENABLE_SALES_SUPERTAG_COLUMN`

### Output

- `GENERATE_INDIVIDUAL_REPORTS`
- `GENERATE_MERGED_MASTER`
- `ENABLE_OUTPUT_FILE_SEPARATED`
- `MAX_ROWS_PER_SHEET` — retained for intermediate preserved-engine batch workbooks; version 1.4.3 does not use it to split the final combined master.

### Salesperson and party scope

- `TEST_MODE`
- `TEST_SALESPEOPLE`
- `TEST_MODE_PARTY`
- `TEST_PARTIES`

The UI supports single or multiple selections, comma/semicolon/new-line custom entries, and an optional low-memory scan of uploaded profile and filter files. For an all-salespeople cloud run, the GitHub worker performs discovery automatically.

### Status, styling and cleanup

- `MATCH_STATUS_FILTER_MODE`
- `MATCH_STATUS_ALLOWED_TYPES`
- `MATCH_STATUS_SPECIFIC_PARTIES`
- `ENABLE_STYLING`
- `ENABLE_NEW_FILL_LOGIC`
- `ENABLE_CLEAN_FILE`
- `ENABLE_PAIR_MODE`
- `ENABLE_TOP_CUSTOMERS_HIGHLIGHT`
- `TOP_CUSTOMERS_N`
- `ENABLE_SMALL_OUTPUT_DELETE`
- `SMALL_OUTPUT_THRESHOLD`

### Advanced sorting

- `ENABLE_ADVANCED_SORTING`
- `SORT_COUNT_THRESHOLD`
- `SORT_NORMAL_PARTY_LIMIT`
- `SORT_HIGH_PRIORITY_PARTY_LIMIT`
- `SORT_EXACT_CAP`
- `SORT_NEAR_CAP`
- `SORT_MAYBE_CAP`
- `SORT_HIGH_PRIORITY_PARTY_TYPES`
- `SORT_HIGH_PRIORITY_TOTAL_HISTORY_THRESHOLD`

## Demand Filter Mapper

All three source controls are available:

- `TOLERANCE_LEVEL`
- `ENABLE_FLO_MAPPING`
- `ENABLE_RELATED_TAG`

## Sort Suggested Output

The app exposes the complete business-control block:

- `ONLY_MATCHED_TYPE`
- `ALLOWED_MATCH_TYPES`
- `COUNT_THRESHOLD`
- Normal/high-priority limits
- Exact/nearby/maybe caps
- High-priority types and history threshold
- Exact/nearby/filter fill RGB sets
- Pair-detection words
- Sales/action type weights and default weight

Input/output paths remain adapter-managed because web uploads use temporary folders rather than Windows or Google Drive paths.

## Separate Excel Sheets

The original source contains path and Windows Excel COM settings rather than matching controls. Those paths are intentionally replaced by the Streamlit upload/download adapter. The source itself remains preserved unchanged.

## GitHub Actions batch wrapper

The cloud wrapper does not remove any runtime setting listed above. The base runtime dictionary is serialized into the private input bundle and restored on the runner.

For each salesperson batch, only these two values are replaced:

```python
batch_runtime["TEST_MODE"] = True
batch_runtime["TEST_SALESPEOPLE"] = current_batch
```

All matching dimensions, tolerance switches, status filters, party restrictions, cleanup, styling, pair protection, output formats, type columns, row limits, and advanced sorting values remain identical across batches.

## Final-output assembly after batching

Batching does not change the requested report formats. After successful batch processing:

- `GENERATE_MERGED_MASTER=True` produces one global `MASTER_CONSOLIDATED_REPORT.xlsx` assembled from all successful batches.
- `GENERATE_INDIVIDUAL_REPORTS=True` places every salesperson workbook together at the final ZIP root.
- `ENABLE_OUTPUT_FILE_SEPARATED=True` places every salesperson's `*_Separated_Parties.zip` together at the final ZIP root.
- Batch logs and runtime snapshots are retained only under `_batch_diagnostics/`.

The assembler aligns batch master columns by header name and recreates the original master styling, decimal formats, and widths with low-memory OOXML streaming. The merged master is written to a single worksheet named `Master_Data` when the total rows fit in one Excel sheet. Report links are written as `HYPERLINK()` formulas so the output does not consume Excel native hyperlink objects.
