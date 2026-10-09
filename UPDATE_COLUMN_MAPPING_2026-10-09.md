# Auto Pricing source-header compatibility hotfix (2026-10-09)

Original master matching/pricing engines are not changed. Only the active Auto Pricing adapter/UI and dataset finalizer are updated. Existing final output names and business-date windows remain unchanged.

## Verified against the uploaded October 9 exports

| Data | Actual RAW processing-date field | Canonical final date field | Date check |
| --- | --- | --- | --- |
| Cart | `Created Datetime` | `CREATED_DATETIME` | 269/269 valid |
| Hold | `Jangad Date` | `JANGAD_DATE` | 27,159/27,159 valid |
| Offer | `Created Date` | `CREATED_DATETIME` | 127/127 valid |
| View | `Appointment Date` | `APPOINTMENT_DATE` | 212/212 valid (individual appointment rows may expand into multiple stones) |
| Wishlist | `Created Date` | `CREATED_DATETIME` | 827/827 valid |
| Sales | `SALE_DATE` | `Sale Date` | 1,516/1,516 valid |
| Stock | none | none | Current snapshot, no date window |

The alias file `processors/auto_pricing_finalizer_v4_1_3/column_mapping_aliases.json` now identifies `Color -> Col`, `Quality -> Qua`, `Created Date -> CREATED_DATETIME` and the renamed inclusions (`TBInc`, `TOIn`, `SBInc`, `SOIn`, `Girdle Open`, `Pav Open`) for Cart, Offer and Wishlist. Old canonical names still take priority. Aliases use exact normalized header names, not fuzzy guesses.

The Auto Pricing "Create Latest Output" page now has a **Check source-to-final column mapping** expander. It shows each final field, resolved source header, and the business date for windowing. Create Auto Pricing is blocked if any required field cannot be mapped. It does not silently substitute one date field for another.

Stock `Expected`: if there is no raw Expected column, a value may be computed from numeric `GRap` and `GPer` using the existing `GRap * ((100 - (GPer + 5))/100) * 0.90` formula. The map shows `DERIVED` only when both operand columns exist; calculations are checked on every row, and invalid rows fail rather than silently passing. Cache-less Expected cells still trigger warnings. This is the existing business formula, not a replacement pricing model.

## Remaining action needed, not automatically repairable

1. **Hold.PARTY_CODE**: the newer `JANGAD_HISTORY_09-10-2026-01-15-55.xlsx` has `Party Name` but lacks Party Code. `Party Name` MUST NOT be used as a code. An earlier 2-Oct Jangad export provides unique code matches for only 19,390 of the 27,159 newer Hold records; 7,769 are not recoverable from that file. Supply a new Jangad export with Party Code or a validated party master crosswalk. Hold remains `FAILED` until safely mapped.
2. **Previous View.xlsx**: failed `Processing_Check.xlsx` identifies an invalid prior `APPOINTMENT_DATE` at source row 23,922. We do not have that prior final workbook, only raw appointment exports. Supply and correct/review the actual previous View.xlsx; do not derive appointment date from Created Date Time.
3. **Previous Sales 120**: failed audit says 11,302 previous dates were invalid. However the separately uploaded `sales_120_file.xlsx` contains 12,236 parseable `Sale Date` values. Confirm this is the intended previous 120-day file for the next run and not a different file from the failed attempt.
4. **KSTAKA.xlsx (RAW Stock)** and **previous sales_180_file.xlsx** were not uploaded. They are needed for full end-to-end reproduction. Other previous action finals are optional but needed to preserve earlier history.

## Validation

- All six supplied current RAW exports have 0 invalid business date values.
- With empty synthetic action histories, the real Cart, Offer, View and Wishlist files all passed their processing stages: Cart 269 output rows, Offer 127, View 2,617 expanded stone rows, Wishlist 827. Hold correctly FAILED because of missing PARTY_CODE; other Hold date parsing succeeded.
- Focused tests (header normalization, date fields, Stock Expected safe derivation and failure, history-date guard, preview, old header preservation, repository) passed: **18 passed**.
- Full repository test collection was not possible in the audit container because `pyarrow` is not installed there; `pyarrow` is included in the app's pinned requirements. The full Streamlit/GitHub production deployment has not been executed.
