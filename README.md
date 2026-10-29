# TINC_NAME + SINC_NAME: short & full names (combined update)

This package COMBINES the two previous updates: **TINC_SINC_COLUMNS_ONLY_UPDATE.zip** and **REGAL_265_STONES_BATCH_FIX.zip**. The existing code already contains two-way normalization; this is a convenient single installation package, **not a new behavior change** over the two fixes together.

## Supported names in BOTH filters and stock

- `MIN` and `MINOR` are the same. E.g., `MIN FEATHER` = `MINOR FEATHER`.
- `MED` and `MEDIUM` are the same. E.g., `MED CLOUD` = `MEDIUM CLOUD`.
- `HVY` and `HEAVY` are the same. E.g., `HVY SPREADED` = `HEAVY SPREADED`.
- `CRYSTEAL 1/2/3` and `CRYSTAL 1/2/3` also normalize to the three user-supplied grade positions.
- `NONE` and `NON` normalize to `NONE`.
- Matching is case-insensitive and ignores surplus whitespace.
- This works in both directions for both `TINC_NAME` and `SINC_NAME`, including inclusive `A TO B` ranges, comma-separated grades, and exclusions beginning with `!`.

The 18-grade order is **the same as the user's original list** (not alphabetical). `MINOR/MEDIUM/HEAVY` are aliases of the 18 grades; they do not introduce extra grades.

## Files included (exact repository paths)

Copy the six files in `original_master_code/` into that directory, and copy `processors/advanced_demand_batch.py` into `processors/`. Five original master files + one NEW `inclusion_rules.py` + one existing batch processor file. No other application files are changed.

**Install all seven files together** if you have not applied both earlier patches; otherwise this package makes no additional logic changes. Back up / commit before replacing, then redeploy Streamlit.

Keep two separate column headers in the `Change Format` worksheet of the active filters workbook: `TINC_NAME` and `SINC_NAME`. Stock file headers need no modifications. If you use an app default filter workbook, ensure it has these updated headers too.

## Verification conducted

- 18/18 grade values tested as short/full equivalents, **both directions**.
- All four main map/filter functions tested for both fields across 6 short/full combinations each.
- For both fields, all 18 grade values produced identical stone IDs in the currently uploaded 1,144-row stock file using short or full filter names.
- REGAL Advanced Demand batch candidate index vs full filter, strict tolerance zero, no other exclusions: **356/356** eligible stones, reproducing the earlier missing-265 fix.
- Focused existing tests: **38 passed, 2 deselected**; the two deselected tests rely on the earlier 18,849-row Stock Actual workbook, but the current `Stock Actual.xlsx` upload is now the 1,144-row BID workbook (identical to `Stock Actual_BID280.xlsx`). This was a test-fixture count mismatch, not a production matching failure.

## Important scope

- This patch does not change unrelated filtering, sorting, prices, Dealmaker, Layout Stock Matcher, `Removed Party`, app defaults, or output formatting.
- Basic/Advanced Demand retain their existing tolerance settings; Suggested/Pricing/BID remain exact for these inclusion columns.
- The 356 eligible stone count applies only to the supplied `regal.xlsx`, 1,144-row stock, tolerance=0, and **no additional exclusions**. Your final app report can differ if `Removed Party`, `no_need.xlsx`, or other controls apply.
