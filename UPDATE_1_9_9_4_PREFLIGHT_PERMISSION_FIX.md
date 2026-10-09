# Auto Pricing v1.9.9.4: Streamlit Cloud preflight permission fix

## Symptom

All uploaded files are Ready, but the dynamic header mapping panel displays:

    Could not check mapping: [Errno 13] Permission denied: 'an internal file'

The Create Auto Pricing button is then blocked by a generic preflight error.

## Root cause

`processors/auto_pricing_finalizer_v4_1_3/process_datasets_v4_1_3.py`
was creating its default output folder (`/mnt/data/output_final`) at import time.
The Streamlit preflight imports that processor merely to inspect mapping rules.
Streamlit Community Cloud may not permit creating that directory.

## Fix

1. Importing processor mapping functions is read-only; output directory creation
   happens only inside `main()` during an actual dataset-processing subprocess.
2. Header preflight reads Sales column names directly from the uploaded XLSX bytes
   in memory, rather than creating a temporary Sales workbook.
3. Normal full processing keeps its existing isolated writable directory, established
   by `run_auto_pricing_full_pipeline` through the `OUTPUT_DIR` environment variable.
4. All required/optional mappings, date windows, Hold PARTY_CODE warning-only rule,
   existing output schemas and business formulas are unchanged.

## Validation

- 43 Auto Pricing regression tests passed (including a regression for imports with
  an unwritable output directory and a no-temporary-directory Sales preflight).
- Current real-world Stock, Sales, Cart, Hold, Offer, View and Wishlist exports,
  plus previous files, passed header preflight with OUTPUT_DIR deliberately set
  to an unwritable directory. 147 fields; 0 blocking missing; the sole optional
  unmapped field was Hold.PARTY_CODE.
- Full-project pytest collection in the current inspection environment is blocked
  by missing optional `pyarrow` dependency. The entire deployed pricing build was
  not exercised against production-sized uploads in Streamlit Cloud.

## Deploy to the existing Streamlit app

Upload/extract the updated project to the GitHub repository connected to your
Streamlit deployment. If you already have v1.9.9.3 deployed, you only need to
replace these project-relative files:

    adapters/auto_pricing_pipeline_adapter.py
    processors/auto_pricing_finalizer_v4_1_3/process_datasets_v4_1_3.py
    VERSION.txt

Commit the changes on the deployed branch, let Streamlit rebuild, then reopen
Auto Pricing and check the mapping preview. No party code master is needed.
If the live app remains unchanged after committing, use the Streamlit app
management controls to reboot it. A browser refresh by itself will not update
server-side code.
