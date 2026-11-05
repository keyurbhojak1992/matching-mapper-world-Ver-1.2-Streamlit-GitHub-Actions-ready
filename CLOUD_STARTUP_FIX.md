# Cloud startup diagnostics and optimization

## What the 2026-11-04 logs prove

- All requirements were installed successfully (including building pandas 2.2.3).
- Community Cloud was using Python 3.14.10.
- Uvicorn started on port 8501.
- Cloud displayed the warning: the app has been CPU-throttled. This is not a Python exception.
- The logs supplied ended at server startup, so they cannot show why the browser had a persistent spinner.

## Changes included here

- In `app.py`, defer selected Excel/Parquet processing adapter imports until the corresponding operation is invoked. No matching algorithms or reports have been altered.
- In `ui/config_controls.py`, import PyArrow only when examining uploaded Parquet metadata instead of when the Overview page loads.
- Add `diagnostic_app.py`: a minimal health page to distinguish frontend/cloud problems from application-code startup.
- Keep the original `requirements.txt`, `.streamlit/config.toml`, file-processing modules, data files, and `.github` workflows unchanged.

## How to deploy

1. Back up your existing GitHub branch/repository and any Streamlit secrets.
2. Copy the changed `app.py`, `ui/config_controls.py`, and optional `diagnostic_app.py` into the **repository root and matching subdirectory**, then commit and push. Alternatively upload this ZIP's project contents into a clean repository, keeping the folder structure.
3. In Streamlit Community Cloud choose **Python 3.12** under Advanced settings when creating a new app. The runtime.txt file alone does not change an already-deployed app's Python version. To change Python on an existing app, Streamlit says you must delete and redeploy it.
4. Select `app.py` as the main file. If it stays blank, try Ctrl+Shift+R or a private/incognito browser window.
5. For an independent check, deploy a second temporary app from the same repo with `diagnostic_app.py` as main file (choose Python 3.12). If **both apps** remain stuck, suspect Community Cloud, browser connectivity or WebSocket problems rather than the matching code. If the diagnostic loads but `app.py` fails, inspect fresh logs and browser DevTools (Network > WS).
6. If the banner says the app has been CPU-throttled, check the indicated expiration time. Python 3.12 should avoid the slow pandas 2.2.3 source build seen in the Python 3.14 logs, which may reduce startup resource pressure. It cannot bypass Streamlit's resource limits.

## Testing done here

All project Python files were compiled for syntax and source imports/entrypoints were checked statically. This environment does not have Streamlit or PyArrow installed, so **a live deployment or full interactive end-to-end test was not performed**. The optimization cannot guarantee a fix for provider-side throttling or a browser WebSocket fault.
