# Matching Mapper World 1.6.0 update

## Implemented

1. Added **Suggested List Splitter** as a separate Streamlit workspace.
   - Recommended section name: `Suggested List Splitter`.
   - Output categories: `Under 10 Stones` and `10+ Stones`.
   - Exactly 10 stones is placed in `10+ Stones`, matching the supplied Colab rule.
   - One nested ZIP is created per salesperson.
   - Files without a recognized Serial No column are preserved in `Needs Review` instead of being silently dropped.
   - Single-salesperson ZIPs retain the salesperson name correctly.

2. Verified the attached **Sort Suggested Output** logic against the app master.
   - The business logic is already equivalent; the attachment only changes the Colab input path/default source location.
   - The Streamlit adapter overrides Colab paths with temporary upload/output paths.
   - Added regression coverage for the strict navigation threshold and Exact/Near by/Filter-only mode.

## Validation performed

- New splitter boundary checks: 9 -> Under 10; 10 and 11 -> 10+.
- Single-salesperson folder detection check.
- Sorting check: Count 200 remains untouched at threshold 200; Count 201 is processed.
- Sorting check: Exact, Near by, and Filter are retained while May be is removed when those are the allowed types.
- Existing repository/UI checks available in the sandbox: 9 passed.
- Full dependency-based suite could not be executed in the sandbox because Streamlit and PyArrow are not installed and outbound package installation is unavailable.
