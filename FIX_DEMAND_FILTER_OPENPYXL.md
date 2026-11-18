# Excel dependency hotfix: Demand Filter Mapper & Advanced Demand Filter

## What the screenshot means

The app reported `No module named 'openpyxl'`. Both of these report processors
need the `openpyxl` Python package to read/write Excel files. The original
uploaded ZIP already pinned `openpyxl==3.1.5` in root `requirements.txt`.
Therefore the reported failure points to the running deployment having a
missing/incomplete Python installation or using a different dependency file.
No stock-matching rule or report algorithm was changed by this hotfix.

## Changes included in the patch

- `requirements.txt`: an explanatory line immediately above the existing
  `openpyxl==3.1.5` pin. Committing this changed dependency file prompts
  Streamlit Community Cloud to refresh its installation.
- `app.py`: both Demand Filter Mapper and Advanced Demand Filter now check
  that Excel support imports successfully before starting any processing.
  If absent, the app displays concrete deployment repair steps.
- `ui/excel_dependencies.py`: small isolated check for `openpyxl`.
- `tests/test_excel_dependency_preflight.py`: regression tests.

## How to fix your deployed app (Streamlit Community Cloud)

1. Back up your GitHub repository branch and any work you need from active
   Advanced Demand Filter checkpoint jobs. Reboots/redeployments may remove
   temporary batch progress.
2. From `PATCH` ZIP, copy each file into the exact same relative GitHub path.
   In particular, do NOT place `requirements.txt` under a nested subdirectory
   unless `app.py` is also there. The correct paths are:

       app.py
       requirements.txt
       ui/excel_dependencies.py
       tests/test_excel_dependency_preflight.py

3. Commit the changes to the branch Streamlit is running.
4. In Streamlit Community Cloud, use **Manage app > menu > Reboot app**, or
   reboot from the workspace. Examine the deployment logs to confirm that
   `openpyxl==3.1.5` installs successfully.
5. Confirm the deployed entrypoint is this repository's `app.py`. If another
   dependency definition exists at the entrypoint, check which takes precedence.
6. Open each demand-filter workspace. If Excel support is still unavailable,
   the new message explains how to correct the deployment instead of starting
   an invalid batch. Retry with new inputs after the environment is repaired.

## Local installation (if running on your own server)

Run these commands from the folder containing `app.py`:

    python -m pip install -r requirements.txt
    python -c "import openpyxl; print(openpyxl.__version__)"
    python -m streamlit run app.py

The same `python` interpreter must be used for `pip` and Streamlit. If using a
virtual environment, activate it before running these commands.

## Verification and limitations

Automated tests and compilation validate the preflight. A successful cloud
reinstall cannot be validated from this offline package; the deployment logs
and actual running app are authoritative. This hotfix does not auto-install
`openpyxl` inside a user request, since that is unreliable in hosted apps.
No original matching processors, data, default rules, or output layout were
changed. A new Advanced Demand Filter run is preferable after redeployment,
as an unfinished job's temporary checkpoint may not survive a reboot.
