@echo off
REM Double-click to open the Confirmation Desk in your browser (Windows).
cd /d "%~dp0"
if not exist .venv\Scripts\streamlit.exe (
  echo First run: setting up, this takes a minute...
  py -3 -m venv .venv || python -m venv .venv
  .venv\Scripts\pip install -q -r requirements.txt
)
where tesseract >nul 2>nul || echo Note: Tesseract is not installed - scanned PDFs will need to be typed in. See README.
.venv\Scripts\streamlit run desk_app.py
pause
