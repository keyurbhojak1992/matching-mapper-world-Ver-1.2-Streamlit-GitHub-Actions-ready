> **Superseded by v1.9.9.3 for optional columns and Hold party codes:** See `UPDATE_OPTIONAL_PRICING_HEADERS_2026-10-09.md`. Missing Hold `PARTY_CODE` no longer blocks pricing. The information below documents the earlier v1.9.9.2 behavior.

# Matching Mapper World v1.9.9.2 — Dynamic header mapping and historical-date fixes

## Main change: dynamic but safe

The **Create Latest Auto Pricing** workspace now includes an interactive **Check / approve source-to-final column mapping** expander. Upload the current sources and previous histories. The preview shows the canonical output field, source header, business-date role, status, match method, and suggestions.

- Known canonical fields and previously verified aliases map automatically. Normalization ignores case, spaces, and punctuation. Existing output field names do not change.
- For any new unknown header, the operator selects a **final standardized column**, selects the actual raw header and clicks **Approve this mapping**. An unapproved fuzzy suggestion NEVER becomes a business field automatically.
- Approved mappings are stored in the current app session, and can be **downloaded as approved_column_mappings.json** and imported in future sessions.
- Manual approvals are passed into the isolated processing subprocess for the current run. Mapping conflicts, missing required sources and date parsing errors block success.
- Do not use a generic helper `Date` as Hold JANGAD_DATE. Sale Date is not the same as Created Date or Invoice Date without explicit user verification.

## Verification of new input files

- **KSTAKA 09102026.xlsx** has `EXP AMT` (verified actual values), mapped to canonical `Expected`. All 15,796 stock output records passed the Expected numeric validation in the individual integration run.
- Previous **View.xlsx** row 23,922 contains no business fields; only an obsolete helper `Date` value remains. Such helper-only rows can be ignored and audited; other invalid history dates still block processing. View produced 26,574 rows in the individual integration run.
- Previous **sales_120_file.xlsx** has 11,302 formula-backed Sale Date values, e.g. `=INT(P2)`, with valid Excel cached dates. Previous Sales reads now use `data_only=True`; cache-less or nonparseable values still fail. Both Sales windows passed the individual integration run.
- **Hold.xlsx** previous finalized history includes verified party-name-to-party-code pairs. Where one name maps uniquely to one code, it can be filled for current Jangad rows; users may also upload an Excel party master having `Party Name` and `PARTY_CODE`. Ambiguous/unmatched names cannot be auto-filled.
- For the supplied files dated October 9, 59 party names (811 relevant Jangad source rows) require confirmed codes. The supplied `Hold_Party_Code_Master_To_Complete.xlsx` is the editable list. The Hold processor correctly refuses PASS until these unresolved records are resolved. Hold output is NOT suitable for production before that.

## Date/window safety

The 90-day Cart/Hold/Offer/View/Wishlist and 120/180-day Sales business windows and boundary refresh rules are unchanged. Stock is a snapshot and has no business date. Every required field must map to the existing canonical output schema, and failed datasets are not delivered as successful outputs.

## Validation and limitations

- 19 targeted tests passed, covering dynamic manual approval, blocking unknown headers, no double-use of a source column, invalid party-name substitutions, safe business dates, cached Sales formula dates, and legacy behavior.
- Individual real-data processing passed for Cart, Offer, Wishlist, View, Stock, Sales 120, and Sales 180. Hold failed as designed due to unresolved external party codes.
- A single all-in-one CLI run in this container exceeded the execution limit. The combined final workbook builder and production Streamlit deployment have not been verified end to end here.

## How to resolve the final Hold blocker

1. Open the provided `Hold_Party_Code_Master_To_Complete.xlsx` and enter the real `PARTY_CODE` for each Party Name in its yellow column. Do not invent codes.
2. Upload that completed workbook into the new **Party-code master** file input in Auto Pricing.
3. Check source-to-final mappings, approve any unknown new headers, then create Auto Pricing.
4. Review `Processing_Check.xlsx` and retain a downloaded `approved_column_mappings.json` for later sessions.
