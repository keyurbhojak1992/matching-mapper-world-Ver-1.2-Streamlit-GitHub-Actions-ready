# Dealmaker v8.2.2 — pair completion stable release

## Carat-range alignment patch — 2026-08-08

- Updated the standard `Carat_grp` ranges to:
  `0.30-0.39`, `0.40-0.49`, `0.50-0.59`, `0.60-0.69`,
  `0.70-0.79`, `0.80-0.89`, `0.90-0.99`, `1.00-1.09`,
  `1.10-1.19`, `1.20-1.24`, `1.25-1.49`, `1.50-1.69`,
  `1.70-1.99`, `2.00-2.49`, `2.50-2.99`, `3.00-3.49`,
  `3.50-3.99`, `4.00-4.49`, `4.50-4.99`, `5.00-7.99`,
  `8.00-14.99`, and `15+`.
- Sales/actions and stock recompute `Carat_grp` from numeric `Carat` so legacy
  carat-group labels cannot mix with the revised grouping.
- If numeric `Carat` is unavailable, an existing supplied `Carat_grp` is kept
  as a fallback.
- No country-specific HK/China carat branch exists in this Dealmaker package,
  so no HK/China logic was changed.
- Pair completion, interval logic, allocation, quality controls, discount
  matching, release gates, and workbook output behavior are unchanged.

## Pair allocation behavior

- Selecting either eligible pair member automatically adds the other member to the same party basket.
- Normal ranking or recipe mismatch does not exclude the related partner.
- A hard exclusion on either member removes the complete pair.
- The pair is globally reserved to one party and cannot be split.
- Added partners are marked with `Pair_Added_Flag`, `Pair_Anchor_Serial`, `Selection_Source`, and `Selection_Reason`.

## Candidate archive correction

v8.2.2 fixes the Stage 3 parser failure caused by appending candidate frames with different optional columns under one CSV header. Stage 2 now consolidates all candidate frames for a batch to a union schema and writes the compressed CSV exactly once. This preserves commas, quotes, newlines, and pair-completion audit columns without changing field counts.

Stage 3 also names the affected batch explicitly if a candidate archive is externally damaged or incompletely written.

## Earlier stability correction retained

Zero-accepted-basket runs create reconciled empty CSV/XLSX/ZIP artifacts and proceed through the independent release audit. The production runner writes a full run log and a focused failure diagnostic when a stage exits unsuccessfully.

## Control workbook

The included control workbook contains the validated Shade, Milky, Open, and Blk rank sequences, including `Shade / GREEN` and the revised Milky/Open canonical mappings.
