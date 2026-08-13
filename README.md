# Matching Mapper World — Streamlit + GitHub Actions worker

A private Streamlit interface for diamond matching, demand mapping, suggested-stock generation, output prioritization, BID drop diagnostics, and workbook distribution.

## Recommended deployment

```text
Private GitHub repository
        ↓
Private Streamlit Community Cloud app
        ↓
Temporary private GitHub draft release
        ↓
GitHub Actions batch worker
        ↓
Private downloadable combined ZIP
```

Start with [`START_HERE.md`](START_HERE.md). The full one-time setup, security model, testing sequence, and troubleshooting guide are in [`STREAMLIT_GITHUB_SETUP.md`](STREAMLIT_GITHUB_SETUP.md).

### New in 1.8.0

- **Auto Pricing**: runs the preserved Auto Pricing calculations from Stock, 180-day Sales, 120-day Sales, and View/Hold/Offer/Wishlist/Cart/Media action files. The two original XLSX outputs are packaged together for download.
- **Dealmaker**: runs the complete supplied v8.2.2 production package, including preflight, staged allocation, deal books, and release-gate audit. The verified bundled `Dealmaker_Full_Control_v6.xlsx` is used by default; an updated control workbook can be uploaded per run without changing the bundled base.
- **Layout Stock Matcher**: runs the preserved layout matching engine from Layout Stock, Sales, and Actions with the source defaults and optional source-supported controls.
- **USA Suggested Stock Mapper** remains available from v1.7.0.

## Heavy Suggested Stock Mapper runs

The Suggested page supports two destinations:

- **GitHub Actions full run** — packages the uploaded inputs, sends them to a private GitHub draft release, and processes salesperson batches on a standard private-repository runner.
- **Run inside Streamlit** — retained for small diagnostic tests only.

For cloud full runs, the worker can discover all salesperson names from the current Parquet profiles and Filters workbook. The optional UI scan is needed only when you want uploaded names in the manual selection lists. Every selected configuration is copied to each batch. Only these runtime values are changed per batch:

```python
batch_runtime["TEST_MODE"] = True
batch_runtime["TEST_SALESPEOPLE"] = current_batch
```

Each batch runs in a fresh subprocess to release pandas, PyArrow, openpyxl, and workbook memory before the next group. The workflow also uses one global concurrency group so two heavy Suggested workers do not run simultaneously.

## Final combined output

Batching is only an internal memory-safety mechanism. After all available batches finish, the adapter restores the normal full-run layout:

```text
suggested_reports_final_combined.zip
├── MASTER_CONSOLIDATED_REPORT.xlsx       # one global merged master, when enabled
├── Mr. Example Salesperson.xlsx          # every individual salesperson workbook, when enabled
├── Mr. Example Salesperson_Separated_Parties.zip
├── ...
├── _batch_diagnostics/                    # logs/runtime snapshots only
├── batch_manifest.json
├── processing_log.txt
└── FINAL_OUTPUT_README.txt
```

The global master is assembled in low-memory streaming mode and preserves the original master output rules for column order, row fills, borders, decimal formats, and widths. Version 1.5.0 strictly creates one worksheet named `Master_Data` whenever the result fits Excel's hard worksheet row limit. Links are written as `HYPERLINK()` formulas, and the intermediate `MAX_ROWS_PER_SHEET` configuration no longer splits the final downloaded master. The preserved source scripts in `original_master_code/` are invoked through adapters; the sorting source now includes hyperlink rebinding and formula translation, and the BID drop diagnostic source has been added as its own process.

## Master-code protection

The verified process sources remain in `original_master_code/` and are invoked through thin deployment adapters:

```text
Suggested_Process.py
Filter_Demand_Stock_Process.py
Separated_Excel_Files.py
Sorting_Suggested_Output_Zip_Process.py
BID_Drop_Checking_Process.py
```

Adapters provide temporary paths, runtime overrides, Linux-compatible upload/download handling, logs, and batching. `SOURCE_CHECKSUMS.sha256` records the included source hashes.

## Other workflows

- **Auto Pricing** — preserves the supplied pricing and cross-verification calculations; only desktop input/output paths are adapted to temporary app files.
- **Dealmaker** — invokes the supplied v8.2.2 production runner directly. Use the bundled control base or upload an updated `Dealmaker_Full_Control_v6.xlsx` for the current run.
- **Layout Stock Matcher** — invokes the preserved layout matching source with Layout Stock, Sales, and Actions uploads.
- **Demand Filter Mapper** — applies the preserved demand/filter process.
- **Sort Suggested Output** — runs the hyperlink-safe sorting process with web-configurable controls.
- **Suggested List Splitter** — groups each salesperson's suggested workbooks into `Under 10 Stones` and `10+ Stones` ZIP folders by counting populated Serial No rows.
- **BID Drop Checking Process** — creates a deep-gap diagnostic workbook from suggested results, bids, sales/actions, stock, filters, and exclusions.
- **Separate Excel Sheets** — uses a Linux-compatible openpyxl adapter because the preserved desktop script depends on Windows Excel COM.

## Repository default data files

The Suggested Stock Mapper and BID Drop Checking Process use repository defaults when the matching upload is omitted:

```text
data/PARTY_PROFILE_PERM.parquet
data/PARTY_PROFILE_TEMP.parquet
data/filters - EBID.xlsx
data/no_need.xlsx
```

Uploaded files always override the repository copy. The four business-data files must be committed to the private repository; the app does not generate them.

## Local development

```bash
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
streamlit run app.py
```

For local GitHub job submission, copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml` and insert the private repository settings. Never commit that file.

## Repository layout

```text
.
├── app.py
├── adapters/
├── cloud_jobs/                   # Streamlit-side bundle and GitHub API client
├── cloud_worker/                 # GitHub Actions runner entry points
├── data/                         # administrator-managed default profiles/rule workbooks
├── .github/workflows/
├── original_master_code/
├── tests/
├── ui/
├── START_HERE.md
├── STREAMLIT_GITHUB_SETUP.md
├── requirements.txt
└── requirements-worker.txt
```

## Data handling

The four administrator-managed default inputs are stored in the private repository under `data/`. Job-specific uploads are temporarily stored as assets on a private draft release and on an ephemeral GitHub Actions runner. The input asset is deleted after processing; the result remains private until the user downloads it and clicks **Delete job files**. Confirm that both repository storage and temporary release storage are permitted by your organization's data policy.

## Copyright

Developed by **Keyur Bhojak** for diamond matching and stock recommendation analysis.

Copyright © 2026 Keyur Bhojak. All rights reserved.
