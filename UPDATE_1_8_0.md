# Update 1.8.0

## New production workspaces

### Auto Pricing

- Added a dedicated **Auto Pricing** workspace.
- Preserved the supplied `Auto Pricing.txt` source as both the original package file and `original_master_code/Auto_Pricing.py`.
- Required inputs: Stock, 180-day Sales, and 120-day Sales.
- Source-defined action inputs are exposed separately: View, Hold, Offer, Wishlist, Cart, and Media.
- The adapter changes only the two hard-coded desktop output destinations. It does not rewrite the 180-day aggregation, 120-day weighted-rate, action-count, cross-action, Diff, or Remark business logic.
- Output ZIP contains the two original business workbooks:
  - `Auto pricing revision.xlsx`
  - `Cross_Verify_Sales_Actions.xlsx`

### Dealmaker

- Added a dedicated **Dealmaker** workspace using the complete supplied **v8.2.2** production package.
- The app invokes `Dealmaker_Production_Runner_v6.py` directly, preserving its existing production sequence:
  1. data preflight,
  2. Stage 1 prepare,
  3. Stage 1B stock/profile preparation,
  4. sequential Stage 2 allocation chunks,
  5. Stage 3 finalization,
  6. Stage 4 deal books,
  7. release-gate audit,
  8. production review ZIP creation.
- Required inputs: full Sales/BID, full Actions, and actual current-stock workbooks.
- The verified bundled `Dealmaker_Full_Control_v6.xlsx` is the default control source.
- A user may download that bundled base, edit it outside the app, and upload an updated `Dealmaker_Full_Control_v6.xlsx` as a **run-only override**.
- The bundled control workbook is never overwritten.
- Original runner-supported sheet-name, header-row, and chunk-size controls are exposed under Advanced settings.

### Layout Stock Matcher

- Added a dedicated **Layout Stock Matcher** workspace.
- Preserved the supplied Layout package and invoke its original `build_outputs()` function directly.
- Required inputs: Layout Stock, Sales, and Actions workbooks.
- Preserved source defaults:
  - tolerance = 1,
  - minimum layout percentage = 10%,
  - minimum absolute stones = 5,
  - minimum action weight = 45,
  - match audit columns enabled,
  - empty salesperson workbooks enabled.
- Source-supported controls are available in the UI without changing their defaults.
- Exact / Nearby / Multi Nearby classification, navigation sheets, salesperson workbooks, and ZIP packaging remain under the preserved source.

## Source protection

- Added the three original uploaded packages under `original_packages/`.
- Added preserved source files/packages under `original_master_code/`.
- Dealmaker's own `SHA256SUMS.txt` is retained and validated in tests.
- `SOURCE_CHECKSUMS.sha256` now covers the preserved source/package assets included with the app.

## Validation

- Auto Pricing fixture verifies sale count, unique buyers, weighted 120-day rate, per-action counts, combined counts, Diff, Remark, and both output workbooks.
- Layout fixture verifies the preserved defaults generate the expected salesperson workbook, party sheet, and five suggested rows.
- Dealmaker validation verifies every file listed in the supplied package checksum manifest plus the bundled control workbook hash, and confirms the original production runner CLI remains intact.
