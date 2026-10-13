# Advanced Demand Filter — Make Config update (v1.9.9.5)

This update applies **only** to the **Advanced Demand Filter** workspace.
The original `original_master_code/Advance_Demand_Filter_Process.py` matching
logic remains unchanged, and all previous matching/settings controls remain available.

## How to use

1. Open **Advanced Demand Filter**, upload the Stock Actual workbook and Filters workbook,
   optionally upload No Need / Clean workbook, and choose the existing matching settings.
2. Click **Make Config** to open the output options.
3. Enter **Minimum stones per party** (whole number, default **0** = disabled).
   - If set to 10, a party with 9 final stones is omitted, and 10 is retained.
   - Counts are based on final, deduplicated party rows **after** No Need and
     output Match Status filters, not on the number of filter rules.
   - Removed parties do not appear in any Excel party sheet or the Navigation
     sheet. A salesperson with no remaining parties is not included in the ZIP.
4. Choose the **Download ZIP layout**:
   - **Salesperson workbooks (current format)** (default): one outer ZIP
     containing `<Salesperson>.xlsx`, with Navigation and one sheet per party.
   - **Separate party files (one ZIP per salesperson)**: one outer ZIP containing
     `<Salesperson>.zip`, each containing `<Party>.xlsx` workbooks. Each party
     workbook contains only its own party sheet; no Navigation links to absent sheets.
   - **Both formats in one ZIP**: one outer ZIP containing both
     `<Salesperson>.xlsx` and `<Salesperson>.zip` for every retained salesperson.
5. Click **Generate advanced demand-filter reports** and download the result ZIP.

## Compatibility and safety

- **0 stones minimum** and **Salesperson workbooks (current format)** reproduce
  the previous default output structure.
- All filtering, style/hyperlink formatting, and matching priority remain unchanged.
- Navigation hyperlinks and counts refer only to retained party sheets.
- The output controls are applied in the report generation stage, not in the
  original demand-matching source or the separate Demand Filter Mapper workspace.
- An empty outer ZIP means no party met the selected conditions; see the
  processing logs for the number of parties below the minimum.

## Testing

`pytest -q tests/test_advanced_demand_filter.py` covers the original default,
threshold boundaries, Navigation, empty salesperson removal, each output layout,
content/format parity, and invalid input controls.
