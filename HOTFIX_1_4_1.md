# Version 1.4.1 deployment hotfix

This package fixes the Streamlit startup error:

```text
ImportError: cannot import name 'LEGACY_RESULT_ASSET_NAME' from 'cloud_jobs.bundle'
```

The error means files from two project versions were mixed in the GitHub repository. In particular, `app.py` was updated while `cloud_jobs/bundle.py` remained from the earlier v1.2 package.

## Recommended update

Replace the complete contents of the GitHub repository with this v1.4.1 project and commit all changed files together.

The most important synchronized files are:

```text
app.py
cloud_jobs/bundle.py
cloud_jobs/github_backend.py
cloud_worker/run_job.py
adapters/suggested_adapter.py
adapters/suggested_result_assembler.py
.github/workflows/process-suggested-cloud.yml
```

Do not copy the enclosing ZIP folder into the repository. The repository root must directly contain `app.py`, `cloud_jobs`, `cloud_worker`, `adapters`, and `.github`.

The Streamlit secret block does not need to be changed.
