@echo off
cd /d "%~dp0"
if exist "dist\ExcelConverter\ExcelConverter.exe" (
  start "" "dist\ExcelConverter\ExcelConverter.exe"
  exit /b
)
if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" run.py
) else (
  py -3 run.py
)
if errorlevel 1 pause
