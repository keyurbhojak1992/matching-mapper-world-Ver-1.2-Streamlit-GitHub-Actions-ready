# Update 1.9.9

## Advanced Demand Filter

- Added a new **Advanced Demand Filter** workspace without changing the existing Demand Filter Mapper.
- Uses the supplied advanced-demand matching source, with requested defaults: tolerance 1, Fancy Color tolerance off, fluorescence mapping off, related Tag_No off, clean/no-need enabled, and output Match Status filtering disabled.
- Requires uploaded **Stock Actual** and **Filters** workbooks.
- **No Need / Clean** is optional. When absent, exclusion logic is skipped; no bundled/default workbook is substituted.
- Missing optional rule sheets are skipped. `Change Format` remains required in the Filters workbook.
- Added streaming large-filter processing in 5,000-rule batches, lazy stock-side candidate indexes, disk-backed duplicate tracking, and disk-backed intermediate match accumulation.
- Preserves source matching/classification, Removed Party behavior, duplicate-rule behavior, optional Tag_No, No Need cleanup, status filtering, sorting, hyperlinks, styles, and salesperson ZIP layout.

## Validation

- Existing Demand Filter Mapper source/adapter restored unchanged.
- Small-workbook parity tests compare the batched output against the supplied source implementation across multiple configuration combinations.
- Optional No Need parity tested.
- Uploaded sample `filters.xlsx` structure processed successfully.
- Synthetic 20,000-rule streaming stress test completed in four 5,000-rule batches.
