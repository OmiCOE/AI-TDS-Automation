@echo off
rem ---- TDS Automation: one-click local launcher (no admin rights needed) ----
cd /d "%~dp0"
where python >nul 2>nul
if errorlevel 1 (
  echo Python was not found. Install it from https://www.python.org/downloads/
  echo Choose "Install Now" - it installs for your user only, no admin rights required.
  pause
  exit /b 1
)
if not exist ".venv" python -m venv .venv
call ".venv\Scripts\activate.bat"
pip install -q -r requirements.txt
echo.
echo Starting TDS Automation - your browser will open. Close this window to stop it.
streamlit run app.py
pause
