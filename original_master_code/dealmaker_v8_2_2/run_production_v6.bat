@echo off
setlocal
if "%~3"=="" (
  echo Usage: %~nx0 SALES ACTIONS STOCK [CONTROL] [OUTPUT_DIR]
  exit /b 2
)
set "CONTROL=%~dp0Dealmaker_Full_Control_v6.xlsx"
if not "%~4"=="" set "CONTROL=%~4"
set "OUTPUT=%~dp0dealmaker_v8_2_2_output"
if not "%~5"=="" set "OUTPUT=%~5"
python "%~dp0Dealmaker_Production_Runner_v6.py" --sales "%~1" --actions "%~2" --stock "%~3" --control "%CONTROL%" --sales-header-row 0 --actions-header-row 0 --stock-header-row 0 --output-dir "%OUTPUT%"
exit /b %errorlevel%
