# Update 1.9.1 — Auto Pricing individual upload workflow

- Full Pipeline now uses individual XLSX upload cards instead of RAW/Input ZIP upload boxes.
- Previous finalized base uploads mirror the familiar Stock / Sales 180 / Sales 120 / action-file layout.
- Latest RAW exports are uploaded individually and are internally assigned canonical dataset names.
- Auto Pricing template upload is removed from the UI; the app uses `data/auto_pricing_template.xlsx` from Git automatically.
- Removed Parties continues to use `data/remove_parties.xlsx` from Git by default, with an optional one-run override.
- Dataset Finalizer v4.1.3 business rules are unchanged.
