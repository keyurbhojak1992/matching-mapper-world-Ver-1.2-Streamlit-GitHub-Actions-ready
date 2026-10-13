# Advanced Demand Filter: 200,000+ Row Stability Patch

Scope: `processors/advanced_demand_batch.py` only. The original matching engine
(`original_master_code/Advance_Demand_Filter_Process.py`), UI controls, database
matching precedence, file names, thresholds, and output layout choices are not
changed. There is no effect on the standard Demand Filter Mapper.

## Improvements

1. Full-stock normalization cache is built lazily, only if a rule needs a full
   stock scan. Indexed INCLUDE rules do not allocate it unnecessarily.
2. SQLite receives matched rows from a generator, avoiding a second giant Python
   list for broad matching rules.
3. A bounded (256 entries and 20,000 cached match links) cache reuses matching
   results for rules with the same **matching predicate**. Party, salesperson,
   city, country and related-tag data are still applied separately. INCLUDE and
   EXCLUDE rules do not share results.
4. Salesperson reports representing 20,000 or more stock-to-party links are
   written using openpyxl's disk-backed write-only worksheets. Individual party
   files switch to the same mode when the party has 20,000+ final output rows.
   Normal-size reports keep the original writer. The streaming writer retains
   navigation sheets, hyperlinks, widths, cell values, formats, row coloring,
   hidden output columns, and worksheet dimensions.
5. Existing 5,000-rule transaction batching and temporary SQLite storage remain.

## Verification performed

- All 12 `tests/test_advanced_demand_filter.py` tests passed, including direct
  comparisons of normal and streaming outputs (values, style, links, sheet order,
  file layout) and a metadata-aware predicate-cache test.
- Synthetic workload A: 200,000 input filter rows (199,999 duplicate rules),
  21,000 matched stock rows, spreadsheet ZIP verified. Peak Python-process RSS
  approximately 146 MB in the test environment.
- Synthetic workload B: 200,000 *unique* filter rules whose business matching
  predicate was identical and whose metadata differed, matching one stock row.
  All 200,000 unique rules processed, 199,999 predicate results reused, output
  verified. Peak Python-process RSS approximately 134 MB.

A separate 200,000-row Stock Actual / 200,000-row output test exceeded the
120-second execution limit of this testing environment; no final ZIP was
verified for that case. The filter-row success cases above do not prove that
200,000 STOCK rows or 200,000 completely different, broad business predicates
will complete within production hosting limits.

## Important remaining constraints

- Stock Actual is still loaded into a pandas DataFrame at once. A 200,000-row
  *stock* workbook with many wide text columns could use substantial RAM.
- The Streamlit adapter/UI still loads upload and final ZIP bytes into memory,
  and Streamlit retains the final output for downloading. Very large ZIPs can
  still exceed hosted memory limits.
- Extremely broad unique rules can produce tens of millions of SQLite match
  records and/or extremely large ZIPs. Temporary disk space and execution time
  must be sufficient. If this applies, run in a properly-sized worker (or add a
  dedicated queued-job flow) rather than increasing the rule batch size alone.
- Very large separate-party/both layouts may produce far more Excel data than
  the default salesperson-only layout.
- Input headers, required sheet names, workbook row limits, and all existing
  matching and exclusion rules retain their original constraints.

## Deployment

Replace the project with the patched full ZIP, or copy
`processors/advanced_demand_batch.py` from the patch-only archive into your
existing project. Preserve the rest of the project unchanged. Install the
existing `requirements.txt` in your production environment. For large jobs,
monitor server RAM, CPU, free temporary disk, and elapsed time. Back up your
current deployment before replacement.
