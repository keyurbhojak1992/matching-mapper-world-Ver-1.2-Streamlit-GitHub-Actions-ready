# Update 1.7.0

## USA Suggested Stock Mapper

- Added a dedicated **USA Suggested Stock Mapper** workspace.
- Accepts one Stock Actual XLSX plus one or more USA sales-history XLSX files.
- Preserves the supplied USA matching logic: top historical Super Tags per party, exact stock matches, and adjacent Color / Clarity / Fluorescence suggestions.
- Produces a ZIP grouped by salesperson, with one XLSX per party.
- Adapts Colab-only paths and `tqdm.notebook` progress handling for Streamlit/server execution.
- Automatically maps a likely stock worksheet to temporary `Sheet1` when needed.
- Adds regression tests for exact/nearby matching and missing sales-history validation.

The existing Suggested Stock Mapper, sorter, splitter, demand mapper, BID checking, and sheet separation workflows remain available.
