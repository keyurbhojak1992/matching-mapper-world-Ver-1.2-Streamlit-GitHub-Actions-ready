ADVANCED DEMAND - ORIGINAL EXCEL FILL / REGAL INCLUSION CLASSIFICATION ONLY

This is an incremental patch for the already-installed TINC_SINC_SHORT_FULL_COMBINED_UPDATE.zip.
It replaces ONE file: original_master_code/Advance_Demand_Filter_Process.py.

What changes:
- When a rule uses TINC_NAME or SINC_NAME, an output stone can be Exact (green)
  only if the entire same rule matches at tolerance 0.
- If the stone needs nonzero tolerance, an otherwise-Exact match is Nearby (yellow).
- Original Maybe (white) behavior from Color/Qua/Flo remains unchanged.
- Multiple INCLUDE rules continue to use the best classification: Exact first.
- All original workbook row-wide cell fills, colors, borders, layout and
  hyperlinks are unchanged. The underlying stock selection is unchanged.
- Does not change unrelated parties/filters or EXCLUDE rule behavior.

Installation:
1. Back up the current file in GitHub.
2. Replace original_master_code/Advance_Demand_Filter_Process.py with the
   matching path in this ZIP.
3. Commit and redeploy Streamlit.
4. Rerun Advanced Demand with the same stock, filters, tolerance (1), and exclusions.

Important: This is based on the latest provided combined patch. If that patch
was not previously installed, install it first.
