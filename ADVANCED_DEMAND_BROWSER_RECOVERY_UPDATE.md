# Advanced Demand Filter - Browser Recovery and Server Concurrency Patch

This update addresses the 2026-10-13 Streamlit log where two different Advanced
Demand jobs (`ab77894c` and `ecceaf7d`) ran interleaved, and an upload/session
reset appeared after matching progressed past 155,000 input rows. The log did
**not** establish a Python exception, OOM kill, or a report-generation failure.

## What changed

1. **Resume link in browser URL.** When a checkpointed job is created or
   manually resumed, the Streamlit URL includes `advanced_demand_job=<full job
   token>`. If uploads/session state disappear after a browser reconnect, open
   that same URL and use **Continue batches** to resume from saved SQLite state.
   The recovery opens *paused* to avoid duplicate processing from two tabs.
2. **One heavy checkpoint at a time on the same server.** A nonblocking global
   OS file lock prevents two different job IDs from simultaneously loading and
   processing stock in separate sessions. Instead, the second receives a
   visible busy message and can retry by clicking Continue batches. This
   supplements the existing per-job lock.
3. **Memory diagnostics in server logs.** Completed checkpoint log lines now
   include `peak_rss_mb` (Linux peak RSS since process start), useful for
   detecting growth. The number `salespeople=0` *during matching* is expected:
   salesperson report files are generated only after matching finishes.
4. **New-job controls.** Active job processing can't accidentally be replaced
   by another "Generate" click in the same tab. Users can explicitly release
   the current job or start another after downloading.

## What is unchanged

All Advanced Demand matching predicates, de-duplication, ranking, tolerances,
No Need, salesperson/party metadata, report layout, and output file generation
remain the same. The original master code and the matching processor are not
modified. This update is compatible with the earlier resumable batch version.

## Important limits

* This update does not accelerate matching or guarantee completion on free
  Streamlit hosting. Reopening the saved URL can recover only if the server
  retains the job's temporary SQLite / input / report files. After a cloud
  container replacement, temp files may be lost; recovery is then impossible.
* Your URL is now a **bearer token** for the job. Keep the link private.
* If batch performance remains poor, reduce "Filter rows per checkpoint" to
  500 or 1,000 **for a new job**. This does not reduce the total match workload.
  Converting the filter-input reader to a durable indexed staging table is a
  separate performance change, which requires validating exact row parity.
* An unusually broad single rule and very large Excel outputs can still use
  substantial RAM; additional changes or a dedicated processing worker may be
  required for such cases.

## Deploy and test

Replace `app.py` and `adapters/advanced_demand_checkpoint_adapter.py` from
this patch in your GitHub repository, then commit and redeploy. Keep every other
file untouched. Start a new batch job after deployment; copy the URL with the
job ID immediately. Run **one job at a time** and do not start the same input
again when a page reloads. Open the same saved URL and press **Continue batches**.

**Test status:** the original 20 targeted Advanced Demand tests and a new
parallel-job-lock regression test passed in the authoring environment. This is
not equivalent to executing your full 200,000-row workbook on Streamlit Cloud.
