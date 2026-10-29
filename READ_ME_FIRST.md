# Minimal inclusion-column patch

Only `TINC_NAME` and `SINC_NAME` have new filtering behavior. All other filtering, sorting, dashboard, and output code is unchanged.

Files in this package are relative to your GitHub repository root. Copy **all six Python files** from `original_master_code/` into the same paths in your app, preserving filenames. **Five existing files are replacements and one (`inclusion_rules.py`) is new.** Do not copy the whole original app ZIP.

Changed existing files:
- `original_master_code/Suggested_Process.py`
- `original_master_code/Pricing_Suggested_Process.py`
- `original_master_code/Filter_Demand_Stock_Process.py`
- `original_master_code/Advance_Demand_Filter_Process.py`
- `original_master_code/BID_Drop_Checking_Process.py`

New file: `original_master_code/inclusion_rules.py`.

In the **Change Format** worksheet, the two separate column headers must be exactly `TINC_NAME` and `SINC_NAME`. Both uploaded stock workbooks already use those headers. You do not need to rename stock-file columns.

Rules support: single grade, comma-separated grades, `A TO B` inclusive ranges, and `!grade` exclusions. Range and tolerance follow the exact 18-grade order you supplied. For the two columns only, `MINOR`=`MIN`, `MEDIUM`=`MED`, `HEAVY`=`HVY`, and `CRYSTAL 1/2/3`=`CRYSTEAL 1/2/3`. An unrecognized inclusion grade or missing stock column raises an error rather than ignoring that condition.

Only the **Basic Demand Filter** and **Advanced Demand Filter** apply the existing configured tolerance level to inclusion rules. Suggested / Pricing Suggested / BID Drop remain exact for these inclusion columns. All other tolerance behavior is unchanged.

**Not changed:** Dealmaker, Layout Stock Matcher, USA suggested mapping, no_need rules, defaults, sales data, other filtering logic. The `Removed Party` vs `Change Format` priority conflict identified earlier is intentionally not addressed.

### How to update
1. Back up or create a Git branch. Extract this ZIP; upload the five replacement files and the one new file to the exact `original_master_code/` locations in your GitHub repository.
2. Upload your updated `filters.xlsx` in the app. If your app uses repository defaults instead of uploads, update the active default `data/filters - EBID.xlsx` with your renamed `Change Format` headers.
3. Restart/redeploy Streamlit and test rules for both headers against regular stock and BID stock. Do not change any other file.

### Verification
40 focused tests passed (including both uploaded stock files and a synthetic end-to-end BID Drop test). The entire original project test suite could not be collected because `pyarrow` is missing from the test environment; this patch does not add/remove any dependencies. The `filters.xlsx` uploaded earlier still has old headings, so it was not replaced; tests use the two new headings. The other previously reported app issues are **not** changed by this patch.
