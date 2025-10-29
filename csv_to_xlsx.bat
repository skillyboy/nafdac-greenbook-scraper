@echo off
REM Wrapper to run csv_to_xlsx.py with the project's venv python on Windows
SETLOCAL
SET VENV_PY=%~dp0venv\Scripts\python.exe
IF EXIST "%VENV_PY%" (
  "%VENV_PY%" %*
) ELSE (
  echo venv python not found at %VENV_PY% - activate your venv or run the script with the correct python executable
  echo Example: %~dp0venv\Scripts\python.exe csv_to_xlsx.py --input nafdac_greenbook.csv --output nafdac_greenbook.xlsx
)
ENDLOCAL
