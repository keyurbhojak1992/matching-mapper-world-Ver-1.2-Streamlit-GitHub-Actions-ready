# Update 1.9.0 — Auto Pricing Full Pipeline

The Auto Pricing workspace now includes the validated Dataset Finalizer v4.1.3 flow.

## New recommended flow
- Upload the latest RAW export ZIP.
- Upload the previous finalized base ZIP.
- Upload the latest Auto Pricing template workbook.
- Optionally override Sales 120 / Sales 180 history files.
- Removed Parties automatically uses `data/remove_parties.xlsx` from the Git project unless an override workbook is uploaded.
- The app finalizes all eight base datasets, validates `Processing_Audit.xlsx`, and builds the formatted Auto Pricing workbook.

## Locked business rules
- Whole boundary-date replacement when RAW contains the previous final's latest calendar date.
- Hold uses RAW `SERIAL_NO`, never `SEQUENCE_NO`.
- View keeps one source row per SERIAL_NO / STONE_ID value; no comma explosion; blank/0 rows are removed.
- Sales refreshes SPer/SRate on `Serial No + Sale Date` overlap and appends new transactions.
- Sales `Amt = SRate x Carat` and derived fields are written as values.
- Blank Lab/Flo is normalized to `NON`.
- Removed Parties applies to RAW and carried history.
- SALE_PERSON is preserved and ignored for reconciliation.
- Broad duplicate deletion remains disabled; duplicate-looking rows stay audit-visible.

The previous prepared-base Auto Pricing mode remains available in a second tab for compatibility.
