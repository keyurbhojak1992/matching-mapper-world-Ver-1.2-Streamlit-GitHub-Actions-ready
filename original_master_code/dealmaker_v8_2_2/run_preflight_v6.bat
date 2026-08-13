@echo off
setlocal
python "%~dp0Dealmaker_Production_Preflight_v6.py" %*
exit /b %errorlevel%
