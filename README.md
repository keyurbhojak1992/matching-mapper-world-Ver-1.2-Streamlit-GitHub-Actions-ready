# REGAL Imports - Full-rule fill classification and Demand Filter Mapper parity

## Scope
Based on the user's earlier `TINC_SINC_SHORT_FULL_COMBINED_UPDATE.zip` and `ADVANCED_DEMAND_ORIGINAL_FILL_COLOR_FIX.zip`, this incremental change edits ONLY:

- `original_master_code/Filter_Demand_Stock_Process.py`
- `original_master_code/Advance_Demand_Filter_Process.py`

Keep the previously installed inclusion-order helper `original_master_code/inclusion_rules.py` and the Advanced Demand candidate index fix `processors/advanced_demand_batch.py`. No other Python files or source stock files change.

## What changed

1. **Demand Filter Mapper:** `SGP Comment` in filters now maps to stock `SGPComment` and uses flexible contains matching (case, spaces, punctuation insensitive); comma-separated positive terms are OR, `!term` exclusions are AND. Also recognizes KeyToSymbol, Comment and Color Desc as known long-text columns when present.
2. **Demand Filter Mapper:** Excel `Exact`/green now requires the **complete filter row** to match at tolerance 0. Selected stones requiring tolerance are `Nearby`/yellow unless their existing multi-grade deviation rule makes them `Maybe`/white.
3. **Advanced Demand Filter:** Same full-rule Exact/Nearby classification for all supported filter columns, not only rows containing TINC_NAME or SINC_NAME. The fast batch processor imports this classifier without further changes.
4. Entire-row fill colors are unchanged: `CCFFCC` green, `FFE699` yellow, `FFFFFF` white. Existing Excel styling, links, row selection rules (apart from the previously ignored SGP condition in Demand Mapper), and layout are retained.

## Install (incremental package)

1. Back up the deployed repository and current output files.
2. Extract the incremental ZIP.
3. Replace the above two Python files at exactly those paths in GitHub. Do not rename or merge them in a single path.
4. Confirm that `original_master_code/inclusion_rules.py` and `processors/advanced_demand_batch.py` are still the versions from the earlier short/full and missing-265 updates.
5. Commit, redeploy Streamlit, and rerun the Demand Filter Mapper and Advanced Demand Filter with the same stock, the same two-row REGAL rules, and tolerance=1.
6. Avoid using an earlier unfinished checkpoint after deploying the code change; run a new Advanced Demand batch job if your deployment resumes cached selections.

## Independent verification (uploaded files)

- Stock: the uploaded `Stock Actual.xlsx` with 1,144 stones (both uploaded stock files have identical SHA256 checksums in this session).
- Filter: the two-row REGAL `regal.xlsx`, with TINC_NAME and SINC_NAME columns.
- Old uploaded Demand Filter Mapper ZIP: 719 stones, all marked green; three Exclusive Clarity rule matches absent.
- New generated Demand Filter Mapper ZIP: 722 stones, 356 green + 366 yellow + 0 white.
- Corrected Advanced Demand Filter: same 722 serial IDs and classifications.
- Restored serials: 553046, 752224, 963396.
- Tests: 22 targeted automated tests passed. A broader smoke suite could not collect because pyarrow is unavailable in this environment.

**Important limits:** This is not a rewrite of the whole application. Suggested Stock Mapper/Pricing Suggested have a different source-aware Exact/Near-by/Filter color scheme, BID Drop uses bid/exclusion statuses, and Dealmaker/Layout Stock Matcher do not read this filter file. Those features were not recolored. Demand Filter Mapper still uses a fixed column mapping for most nontext fields; it is not yet fully equivalent to Advanced Demand Filter's future-column auto-mapping. For active rules with unrecognized columns, the Demand Mapper's legacy skip behavior remains; audit those independently before production.

## Ready-to-inspect output

- `DEMAND_FILTER_MAPPER_CORRECTED_OUTPUT.zip` - corrected Demand Mapper sample using these two REGAL filter rows.
- `DEMAND_FILTER_MAPPER_CORRECTED_REGAL.xlsx` - its extracted Excel workbook.
- `REGAL_DEMAND_MAPPER_AUDIT.xlsx` - ID-level discrepancies and verification.

The sample uses only the two REGAL rules from `regal.xlsx`, without Removed Party/no_need exclusions. Real output can differ with additional settings and other filters.
