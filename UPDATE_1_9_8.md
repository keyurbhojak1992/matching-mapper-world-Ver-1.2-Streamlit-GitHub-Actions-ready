# Update 1.9.8

- Pricing Last 5 BID now runs a dedicated matching engine aligned with `pricing_team_single_pipeline.py`, including the pricing carat buckets and fluorescence mapping defaults; the regular Suggested Stock engine keeps its existing profile-compatible carat ranges.
- Fixed the Last-5-BID mismatch where the generated pricing profile used narrow carat groups but the app matcher still used older broad groups.
- Default Suggested/BID filter rules are now embedded in `defaults/suggested_rules.py` (1,562 removed parties, 142 Change Format rules, 1 tolerance rule, 260 no-need rules).
- `filters.xlsx` / `no_need.xlsx` are no longer required repository files. Optional uploads still override the built-in dictionaries.
- Suggested UI defaults now match the attached master-code defaults more closely (scope, report type, status filtering, sorting threshold, and priority types).
