# Matching Mapper World 1.5.0 update

## Implemented

1. **Sort Suggested Output** now uses the supplied hyperlink-safe sorting logic.
   - Standard Excel hyperlinks are copied to every destination cell and rebound to the new coordinate.
   - Formula cells, including `HYPERLINK()` formulas, translate relative references when a row moves.
   - Existing filtering, scoring, limits, workbook order, formatting, comments, and navigation-count behavior remain unchanged.

2. **Suggested Stock Mapper** now treats these uploads as optional overrides and falls back to exact repository paths:
   - `data/PARTY_PROFILE_PERM.parquet`
   - `data/PARTY_PROFILE_TEMP.parquet`
   - `data/filters - EBID.xlsx`
   - `data/no_need.xlsx`

   The same fallback behavior is applied to local Streamlit processing and GitHub Actions full processing.

3. Added **BID Drop Checking Process** as a separate Streamlit workspace.
   - Required uploads: `MASTER_CONSOLIDATED_REPORT.xlsx`, BID workbook, Sales & Actions workbook, and Stock Actual workbook.
   - Optional overrides: Filters and No-need workbooks.
   - Defaults when omitted: `data/filters - EBID.xlsx` and `data/no_need.xlsx`.
   - Output: `Deep_Gap_Analysis_<timestamp>.xlsx`.

## Repository data note

The attached source ZIP did not contain the four binary default data files. The fallback code and exact paths are implemented, but the verified business-data files must be added to the private repository's `data` folder before omitted-upload processing can run.

## Validation performed

- Python compilation of app, adapters, cloud worker, master sources, and tests.
- Source checksum verification.
- Sorting regression test for row movement, standard hyperlink rebinding, and formula translation.
- BID Drop Checking smoke test, including repository-default Filters and No-need resolution.
- Existing repository/UI tests and non-Parquet processor/cloud tests.
