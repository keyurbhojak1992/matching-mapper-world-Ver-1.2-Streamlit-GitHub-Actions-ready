Dataset Finalizer + Auto Pricing Output v4.1.3

FINALIZED RULES
1. Production date is automatic Asia/Kolkata current date. TEST_PROCESSING_DATE / --test-date exists only for regression.
2. Cart / Offer / Wishlist / Hold:
   - read previous final history;
   - identify the previous final's LAST CALENDAR DATE;
   - when current RAW contains that same date, DELETE that whole date from previous history and rebuild the whole date from RAW;
   - append later RAW dates;
   - if RAW does not contain the boundary date, append only records newer than the previous final's latest timestamp.
3. Hold SERIAL_NO always maps from RAW SERIAL_NO. SEQUENCE_NO is activity sequence and must never replace the stone serial. Digit length is not used to decide mapping.
4. View:
   - same whole-boundary-date refresh principle;
   - one source row stays one row;
   - STONE_ID/SERIAL_NO is NOT comma-split/exploded;
   - blank/0 SERIAL_NO rows are removed;
   - 90-day cutoff uses processing date.
5. Sales 120/180:
   - supplied previous Sales files remain descriptive history;
   - transaction identity = Serial No + Sale Date;
   - exact RAW overlap updates SPer and SRate only (Amt is recalculated as SRate x Carat);
   - new RAW transactions append;
   - removal-party filtering applies to RAW and carried history.
6. Blank Lab or Flo becomes NON.
7. Stock is current RAW snapshot. Serial-only export footer/non-data rows are removed.
8. SALE_PERSON is preserved as source/history data and is not used as a reconciliation criterion.
9. Exact row deduplication is still deferred. v4.1.3 audits duplicates but does not delete them.
10. Auto Pricing formatted workbook is rebuilt from the processed bases. Sheet1 metrics are recalculated from current bases (not copied from old static values):
    - Sold_Avg_Rate_120days = SUM(Amt) / SUM(Carat) by Tag from Sales 120
    - Diff = ROUND((Expected - Sold_Avg_Rate_120days) / Expected * 100, 2), or 0 when Expected=0
    - Remark = Yes only when sold average > 0 and Diff > 0; otherwise No
    - Unique_Buyer_Count / Sold_Party / sale_count come from Sales 180 by Tag
    - ALL_Action_count = distinct action parties for that exact Serial No across Hold/Wishlist/View/Offer/Cart
    - Action_count = sold buyers for the Tag that also appear in that stone's action-party set

FILES
- process_datasets_v4_1_3.py : validated base-data processor
- build_auto_pricing_v4_1_3.py : formatted Auto Pricing workbook builder
- run_all_v4_1_3.py : end-to-end ZIP runner
- column_mapping_aliases.json : editable source-header aliases

PRODUCTION EXAMPLE
python run_all_v4_1_3.py \
  --raw-zip Raw.zip \
  --input-zip "Input latest.zip" \
  --template "Latest Auto pricing.xlsx" \
  --remove-party "Remove party.xlsx" \
  --output-dir output

28-AUG-2026 REGRESSION EXAMPLE
python run_all_v4_1_3.py \
  --raw-zip Raw.zip \
  --input-zip "Input 28-08-2026.zip" \
  --sales-120-history sales_120_file_input.xlsx \
  --sales-180-history sales_180_file_input.xlsx \
  --template "Auto pricing revision 28-08-2026_Without Formula OLD.xlsx" \
  --remove-party "Remove party.xlsx" \
  --test-date 2026-08-28 \
  --output-dir output_test

IMPORTANT
Do not set --test-date for live/production runs. It is only to reproduce historical regression tests.
