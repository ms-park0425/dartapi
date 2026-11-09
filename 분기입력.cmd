@echo off
chcp 65001 >nul
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" "quarter_picker.py"
) else (
  py -3 "quarter_picker.py"
)
if errorlevel 1 pause
