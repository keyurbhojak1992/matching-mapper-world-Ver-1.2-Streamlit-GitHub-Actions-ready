# Matching Mapper World v1.9.9.3 — Optional pricing-column policy

This is a revision to **v1.9.9.2**. The canonical output schemas and pricing formulas are unchanged. The new preflight checks whether a field is actually required by Auto Pricing, by the 90/120/180-day date-window rules, or by the stone/party matching logic. Only these required fields may block processing when their source headers are missing.

## Required input categories

- **Cart, Offer, Wishlist:** `CREATED_DATETIME`, `Serial No`, `Sold Party`.
- **Hold:** `JANGAD_DATE`, `SERIAL_NO`, `Sold Party`. `PARTY_CODE` is optional.
- **View:** `APPOINTMENT_DATE`, `SERIAL_NO`, `Sold Party`.
- **Sales:** `Serial No`, `Shp`, `Carat`, `Col`, `Qua`, `Lab`, `Flo`, `Sold Party`, `Sale Date`, `SRate`.
- **Stock:** `Serial No`, `Shp`, `Carat`, `Color name`, `Qua`, `Lab`, `Flo`, plus `Expected` (directly supplied or computed from both `GRap` and `GPer`). If `Expected` is supplied, the `GRap` and `GPer` source columns are optional.

All non-listed **output** columns are retained in the same position and may be mapped when available. Unmapped optional fields are blank in new records and shown in `Processing_Check.xlsx` as `OPTIONAL_MISSING` / `WARNING`, instead of causing failure. The dynamic mapper still offers approved overrides for new header names; it never silently guesses a date, a party code, or two targets from one source column.

## `PARTY_CODE` and Hold history

The Hold `PARTY_CODE` field **does not participate in the pricing calculation**. The code is populated only when the source provides it or a single verified party-name/code match exists in the optional previous Hold history / optional party master. Unverified and ambiguous codes remain blank. The output retains the `PARTY_CODE` column, and the Processing Check reports the number of missing codes as a WARNING rather than an ERROR. A party-code master is no longer a prerequisite to pricing.

## Data integrity safeguards still blocking

- A required pricing field, stone identifier, party name or authoritative business-date header cannot be mapped.
- Current source rows contain invalid required dates or previous-history rows have invalid authoritative dates.
- Sales raw-to-final mapped values are unexpectedly lost, Sales pricing rate is missing, or Stock Expected cannot be validated or derived numerically.

## Validation (2026-10-09)

- 40 focused regression tests passed, including safe handling of missing optional action attributes, Sales `Cut/Pol/Sym/SPer`, Stock `GRap/GPer` when Expected exists, Hold missing `PARTY_CODE`, and failures for missing business date / pricing rate.
- Individual real-file Hold processing: **PASS**, 90,561 retained output rows, zero ERROR-level issues; `PARTY_CODE` was unavailable in the new Jangad source and was populated only where verified historically.
- The seven current source exports were checked at the header level: the only unmapped field is `Hold.PARTY_CODE`, now optional.
- The single combined full-data processing attempt exceeded the execution limit, so an end-to-end production workbook-build result is not claimed.

`UPDATE_DYNAMIC_HEADER_MAPPING_2026-10-09.md` describes the earlier v1.9.9.2 stricter behavior; the policy in this file supersedes its Hold-code blocking requirement.
