Dataset Finalizer + Auto Pricing Output v4.1.7

FINALIZED RULES
1. Production date is automatic Asia/Kolkata current date. TEST_PROCESSING_DATE / --test-date exists only for regression.
2. Cart / Offer / Wishlist / Hold:
   - read previous final history;
   - identify the previous final's LAST CALENDAR DATE;
   - when current RAW contains that same date, delete that whole date from previous history and rebuild the whole date from RAW;
   - append later RAW dates;
   - if RAW does not contain the boundary date, append only records newer than the previous final's latest timestamp.
3. Hold SERIAL_NO always maps from RAW SERIAL_NO. SEQUENCE_NO is activity sequence and must never replace the stone serial.
4. View:
   - boundary/window date is APPOINTMENT_DATE;
   - same whole-boundary-date refresh principle;
   - comma/semicolon-separated STONE_ID values are exploded to one SERIAL_NO row per stone;
   - blank/0 stone IDs are removed;
   - 90-day cutoff uses processing date versus APPOINTMENT_DATE;
   - exact full-row duplicates are removed.
5. Sales 120/180:
   - previous Sales files are descriptive history before their latest Sale Date;
   - if current RAW contains that latest calendar date, replace the entire boundary date from RAW and append later RAW dates;
   - if RAW does not contain that boundary date, append only rows newer than the previous file's latest timestamp;
   - SPer is optional in current RAW and is not used by Auto Pricing metrics;
   - exact full-row duplicates are removed.
6. Blank Lab or Flo becomes NON. This includes blank Stock Lab -> NON before Tag creation.
7. Stock is current RAW snapshot. Serial-only footer/non-data rows are removed. Formula-backed Expected values are materialized to numeric values from the source workbook cache, with a GRap/GPer fallback if the cache is unavailable.
8. SALE_PERSON is preserved as source/history data and is not used as a reconciliation criterion.
9. Auto Pricing formatted workbook is rebuilt from the processed bases. Sheet1 metrics are recalculated from current bases:
   - Sold_Avg_Rate_120days = SUM(Amt) / SUM(Carat) by Tag from Sales 120
   - Diff = ROUND((Expected - Sold_Avg_Rate_120days) / Expected * 100, 2), or 0 when Expected=0
   - Remark = Yes only when sold average > 0 and Diff > 0; otherwise No
   - Unique_Buyer_Count / Sold_Party / sale_count come from Sales 180 by Tag
   - ALL_Action_count = distinct action parties for that exact Serial No across Hold/Wishlist/View/Offer/Cart
   - Action_count = sold buyers for the Tag that also appear in that stone's action-party set

DROP-IN FILE NAMES
The package intentionally keeps the original module/folder names (auto_pricing_finalizer_v4_1_3/...) so existing Streamlit imports do not need to change. The internal processor version is 4.1.7.
