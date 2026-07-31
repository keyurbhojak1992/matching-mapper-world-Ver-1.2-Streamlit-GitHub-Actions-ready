# Streamlit + GitHub Actions deployment guide

## What this version does

The Streamlit app remains the private upload/configuration interface. The heavy Suggested Stock Mapper run is moved to a standard private-repository GitHub Actions Linux runner.

The app stores four administrator-managed default inputs in the private repository under `data/`: the permanent and temporary Parquet profiles, `filters - EBID.xlsx`, and `no_need.xlsx`. User-uploaded Stock Actual, overrides, and job-specific files are not committed; each submitted job uses a temporary **private draft release**, uploads `input_bundle.zip`, and dispatches `.github/workflows/process-suggested-cloud.yml`.

The worker:

1. Downloads and safely extracts the bundle.
2. Discovers salesperson names from the current profiles and Filters workbook when **All salespeople** is selected.
3. Runs each batch in a fresh Python subprocess so memory can be returned to the operating system between batches.
4. Copies the complete selected configuration to every batch.
5. Overrides only `TEST_MODE=True` and `TEST_SALESPEOPLE=current_batch`.
6. Publishes the final ZIP and status files back to the same private draft release.
7. Deletes the uploaded input asset when the workflow finishes.

The verified process sources in `original_master_code/` are called through adapters. The sorting source includes the approved hyperlink-safe update, and the BID drop source is included as a separate diagnostic process.

---

## A. Prepare the private GitHub repository

### A1. Upload the project

Create a **private** GitHub repository, then upload all project files. These paths must exist on the `main` branch:

```text
app.py
requirements.txt
requirements-worker.txt
.github/workflows/process-suggested-cloud.yml
cloud_jobs/
cloud_worker/
adapters/
original_master_code/
ui/
```

Do not upload `.streamlit/secrets.toml`. It is intentionally excluded by `.gitignore`.

### A2. Enable Actions

Open:

```text
Repository → Settings → Actions → General
```

Under **Actions permissions**, allow GitHub Actions. Under **Workflow permissions**, choose:

```text
Read and write permissions
```

Save the setting. The workflow also declares `contents: write`, which it needs to read, replace, and delete private release assets.

If the repository belongs to an organization, an organization policy may control these settings. Ask the organization owner when the options are locked.

---

## B. Create the fine-grained token used by Streamlit

The Streamlit server needs a token to create the private draft release and dispatch the workflow. The workflow itself uses GitHub's built-in temporary `GITHUB_TOKEN`; no long-lived token is placed in the workflow.

Open:

```text
GitHub profile picture
→ Settings
→ Developer settings
→ Personal access tokens
→ Fine-grained tokens
→ Generate new token
```

Recommended values:

```text
Token name: matching-mapper-streamlit
Expiration: choose a manageable expiry and renew before it expires
Repository access: Only select repositories
Selected repository: your Matching Mapper private repository
```

Repository permissions:

```text
Actions: Read and write
Contents: Read and write
Metadata: Read-only (automatically included)
```

Copy the token immediately. Store it only in Streamlit secrets or a trusted password manager.

### Never put the token in

```text
app.py
README.md
.gitignore
GitHub commits
screenshots
email or chat messages
```

Revoking the token immediately stops new submissions from the Streamlit app.

---

## C. Deploy to Streamlit Community Cloud

1. Sign in to Streamlit Community Cloud with the GitHub account that administers the repository.
2. Connect Streamlit to private repositories when prompted.
3. Click **Create app**.
4. Select:

```text
Repository: YOUR_USERNAME/YOUR_PRIVATE_REPOSITORY
Branch: main
Main file path: app.py
```

5. Open **Advanced settings**.
6. Select Python `3.12`.
7. Paste this into **Secrets**:

```toml
[github]
repository = "YOUR_USERNAME/YOUR_PRIVATE_REPOSITORY"
token = "github_pat_YOUR_REAL_TOKEN"
workflow = "process-suggested-cloud.yml"
ref = "main"
```

8. Save and deploy.

A copy without a real token is provided at `.streamlit/secrets.toml.example`.

### Keep the app private

Open the deployed app's **Share** menu and choose:

```text
Only specific people can view this app
```

Invite only trusted work email addresses. The app contains a server-side credential capable of creating releases and starting Actions jobs in the selected repository. Streamlit does not expose the secret to viewers, but viewers can submit jobs through the app.

---

## D. First end-to-end test

Do not begin with all salespeople.

1. Open **Suggested Stock Mapper**.
2. Upload Stock Actual. Upload Filters or profile files only when overriding the repository defaults.
3. Click **Scan uploaded/default files for salesperson and party names** only when the required test name is not already in the selection list. This scan is optional for an **All salespeople** cloud run because the worker performs its own discovery.
4. Configure the matching and output controls exactly as needed.
5. In **Salesperson / party scope**, enable salesperson selection and choose one known salesperson.
6. Under **Processing destination**, select:

```text
GitHub Actions full run (recommended)
Batch scope: Only salespeople selected above
Salespeople per batch: 1
Continue after failed batch: ON
```

7. Click **Submit full run to GitHub Actions**.
8. Note the displayed job tag.
9. Click **Refresh status** periodically.
10. When the result is ready, click **Create result download link**, then download it.
11. Open `batch_manifest.json` inside the result and confirm the selected salesperson succeeded.
12. Tick **Allow cleanup** and click **Delete job files**.

You can also inspect the run in:

```text
GitHub repository → Actions → Process Suggested Stock Mapper job
```

---

## E. Full run settings

After the one-person test succeeds:

```text
Batch scope: All salespeople discovered from current inputs
Batch size: 1 initially
Continue after failed batch: ON
```

The worker reads only the salesperson columns for discovery. It then runs the preserved master process separately for each batch.

Increase batch size to `2` and then `3` only after observing successful memory use. Batch size affects how many salesperson outputs are created in one preserved-master run. It does not eliminate the need to load the common Stock, Filters, and profile files for each batch.

The app blocks a new job while another Matching Mapper draft release still reports `UPLOADING`, `QUEUED`, `STARTING`, or `RUNNING`. The workflow also uses a repository-wide concurrency group as a second safeguard against simultaneous heavy workers.

---

## F. Result structure

Batch folders are an internal processing detail. The downloadable result now restores the same root layout produced by a normal unbatched master run:

```text
suggested_reports_final_combined.zip
├── MASTER_CONSOLIDATED_REPORT.xlsx       # when Generate merged master is ON
├── Mr. Aakash Jain.xlsx                  # when individual reports are ON
├── Mr. Aakash Parmar.xlsx
├── ...
├── Mr. Aakash Jain_Separated_Parties.zip # when separated output is ON
├── ...
├── _batch_diagnostics/
│   ├── batch_01/
│   │   ├── processing_log.txt
│   │   ├── batch_runtime.json
│   │   └── worker_console.txt            # when console output exists
│   └── batch_02/
│       └── ...
├── batch_manifest.json
├── processing_log.txt
└── FINAL_OUTPUT_README.txt
```

When **Generate merged master** is enabled, every successful batch master is streamed into one global `MASTER_CONSOLIDATED_REPORT.xlsx`. Columns are aligned by header name, and the adapter recreates the master-code styling, decimal formats, and widths without loading the entire global workbook into RAM. The merged master is written to one worksheet named `Master_Data` when the row count fits in one Excel sheet; links are written as `HYPERLINK()` formulas.

When **Generate individual salesperson reports** is enabled, every salesperson workbook is copied directly to the final ZIP root. When **Create separated party ZIPs** is enabled, all salesperson-specific separated ZIPs are also copied to that same root. There are no nested `batch_XX/suggested_reports.zip` files in new results.

`batch_manifest.json` records successful, failed, and skipped batches and includes a final-output summary. When **Continue after failed batch** is on, successful batch outputs are still combined even if another group fails.

---

## G. Security and data retention

### Stored locations

During a job, business files exist in:

1. The authorized user's browser/Streamlit upload session.
2. A private draft release asset in the private GitHub repository.
3. The temporary GitHub Actions runner disk.
4. The private result draft release until cleanup.

The input asset is deleted at the end of the workflow. The result remains so it can be downloaded after the runner has disappeared. Delete the job from the Streamlit panel after confirming the download.

### Important policy check

This design temporarily stores business Excel and Parquet files on GitHub infrastructure. Confirm that this is acceptable under your organization's confidentiality, customer-data, and data-residency policies before uploading production data.

### Signed result links

The app first asks GitHub for a short-lived signed download redirect. Treat that link as confidential until it expires. Do not forward it outside the authorized team.

---

## H. Free-runner limits

A standard Linux runner for a private repository currently provides approximately:

```text
2 CPU cores
8 GB RAM
14 GB SSD
```

The workflow timeout in this repository is `350` minutes, just under GitHub's normal six-hour job ceiling.

Both the input bundle and final combined ZIP must fit in a single GitHub release asset. This project keeps a safety margin and rejects files at approximately `1.95 GB`.

Batching reduces simultaneous output/workbook pressure and forces memory cleanup between batches. However, the preserved master engine still loads the shared Stock, Filters, profile data, and optional sales history inside each batch. If those common inputs plus one salesperson require more than 8 GB, even batch size `1` can fail.

Private-repository Actions usage consumes the minutes included with your GitHub plan. Check **Settings → Billing and licensing** before regular daily use. Without available included minutes, jobs will not run until quota is available or billing is configured.

---

## I. Troubleshooting

### “GitHub repository must use OWNER/REPOSITORY format”

Correct the Streamlit secret:

```toml
repository = "username/repository-name"
```

### “The configured GitHub repository is public”

Make the repository private before production file uploads.

### HTTP 401

The fine-grained token is invalid, expired, revoked, or pasted incorrectly.

### HTTP 403

Check that the token is restricted to the correct repository and has:

```text
Actions: Read and write
Contents: Read and write
```

Also confirm that repository Actions and workflow write permissions are enabled.

### Workflow does not start

Confirm that this exact file exists on the configured branch:

```text
.github/workflows/process-suggested-cloud.yml
```

Confirm the secret has:

```toml
workflow = "process-suggested-cloud.yml"
ref = "main"
```

### Job says another job is active

Open **Recover a previous full-run job** in Streamlit. Load the job, inspect its status, and delete it only after it has completed or you have confirmed it is stale.

### Process is killed or reports out-of-memory

Use batch size `1`, disable nonessential output formats for diagnosis, and avoid enabling advanced sorting until the base run succeeds. If one-person processing still exceeds 8 GB, the free GitHub runner is not large enough for that dataset.

### Result is too large to publish

Reduce output options or submit smaller selected-salesperson jobs. A single release asset cannot hold the oversized combined ZIP.

### Streamlit app was restarted and lost the current job panel

Use **Recover a previous full-run job** and either paste the saved job tag or click **Find recent jobs**.

---

## J. Daily operating procedure

```text
1. Open the private Streamlit app.
2. Upload the current input files.
3. Choose all matching and output settings.
4. Submit one GitHub Actions job.
5. Close the page if necessary; the job continues.
6. Return later and recover/refresh the job.
7. Download and verify the final ZIP.
8. Delete the job files.
```
