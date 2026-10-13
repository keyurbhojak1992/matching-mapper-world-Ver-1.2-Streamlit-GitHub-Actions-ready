# Advanced Demand Filter: Checkpointed / resumable batches

## Purpose

A single Streamlit processing call can be interrupted by server CPU/RAM limits,
websocket disruptions, or application restarts. The existing 5,000-rule SQLite
transaction batching was still contained in **one** long call. This update adds
an optional checkpointed workflow, enabled by default on the Advanced Demand
Filter page. The existing normal workflow remains available by switching it off.

## How to use

1. Upload Stock Actual.xlsx and Filters.xlsx, plus optional No Need.xlsx.
2. Select **Safe checkpointed batches**, on by default.
3. Start with **1,000 filter rows per checkpoint**. For especially broad rules or
   lower-memory hosts use 500. For faster jobs when resources permit use 2,000
   or 5,000.
4. Click **Generate advanced demand-filter reports**. The screen shows processed
   source rows, unique rules, report stage, and a **batch job ID**.
5. The app automatically runs the next checkpoint. You can pause/continue or
   enter the batch job ID under **Resume an existing batch job** when reconnecting
   to the **same server** and its temporary data still exists.
6. When the job reaches **Finished**, download the original-format ZIP.

## What is checkpointed

- **Filter matching:** one transaction per checkpoint, with persistent SQLite
  deduplication, party/salesperson metadata, rank merging, and cursor position.
  A failed checkpoint rolls back both the matched rows and the cursor.
- **Time-bounded matching:** checks elapsed time between source rows (default
  about 25 seconds per matching step). An unusually slow *single* rule is not
  preemptible; it can still exceed host limits.
- **Stock loading:** the normalized pandas stock DataFrame is serialized once
  into a temporary pickle and reused each run, preserving its data and dtypes.
- **Report creation:** one salesperson group per checkpoint. Each finished
  salesperson file is staged on disk and tracked atomically in SQLite. Large
  workbooks use the existing write-only Excel writer.
- **Final ZIP:** assembled from the recorded finished artifacts on disk in the
  same original group order and filenames.
- **Log entries:** the Streamlit server console receives job ID, stage, input
  row offset, and salesperson groups on each successful checkpoint.

## Preserved behavior

The original matching engine under `original_master_code/` is unchanged. All
rule predicates, tolerance options, filters, no-need exclusions, rank merging,
related tags, per-party minimum, salesperson/party grouping, Make Config output
choices, output Excel values, sheet navigation, and styles are retained. The
normal non-checkpointed execution remains supported.

Added code files:

- `processors/advanced_demand_incremental.py`
- `adapters/advanced_demand_checkpoint_adapter.py`
- `tests/test_advanced_demand_incremental.py`

Modified code: `app.py` only, to expose and drive the new workflow. The earlier
`processors/advanced_demand_batch.py` 200K Excel streaming patch is still part
of this package and remains unchanged.

## Tests and limitations

- **20 targeted tests passed**, including 12 earlier Advanced Demand tests,
  all three output layouts, cell values, hyperlink/format comparisons, No Need,
  removed parties, duplicate rules across checkpoints, status filters, rollback
  after simulated exception, and time-limited continuation.
- A 10,000-unique-filter-row smoke test completed in seven calls (five matching,
  one report, one ZIP package).
- The complete app suite could not be collected in the authoring environment
  due to the missing unrelated `pyarrow` dependency. The deployment
  requirements were not changed.

**Important limits:** This is not a guarantee that every 200,000+ stock/filter
input succeeds on free Streamlit hosting. A single extremely broad rule, a
single oversized salesperson workbook, or a very large final ZIP/download can
still exhaust server resources. Matching still loads the stock DataFrame into
memory once per checkpoint; processing a single rule is not subdivided.

**Resume durability:** Jobs, SQLite, cache, and staged reports live in the
server's **temporary directory** (automatically cleaned after ~24 hours of
inactivity when a new job starts). Browser reconnections or Python reruns can
resume them while that same server retains its temporary storage. A Streamlit
redeployment, server replacement, container rebuild, disk exhaustion, or host
cleanup may wipe the files. For restart-proof cloud execution, use an external
persistent object store / database and a dedicated worker, which is a separate
infrastructure change.

**Disk and memory:** Your host needs space for uploaded files, SQLite matches,
stock cache, intermediate Excel files, staged group archives, and final ZIP.
The Streamlit download widget also loads the final ZIP for download; very large
outputs can still exceed the hosting memory budget. Pause button response may
wait until the current checkpoint finishes. If no output is created because no
party meets filters, the completed ZIP will be empty, matching the old logic.

## Deploy

Back up your GitHub repository first. Extract the **patch-only ZIP** at the
repository root so `app.py`, `adapters/...`, and `processors/...` paths match.
Commit and redeploy. Or use the full clean source archive. No new requirements
are needed. Do not copy just one Python file; UI, adapter and processor must be
updated together.
