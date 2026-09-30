@echo off
rem One command for graders on Windows:  run_all.bat C:\path\to\data
rem Installs into a local .venv, runs both tasks, the vendor forms and the tests. Outputs go to .\output
setlocal
set "PYTHONUTF8=1"
set "DATA=%~1"
if "%DATA%"=="" set "DATA=C:\data"
set "HERE=%~dp0"
set "OUT=%HERE%output"

set "SYS_PY="
where py >nul 2>nul && set "SYS_PY=py -3"
if not defined SYS_PY where python >nul 2>nul && set "SYS_PY=python"
if not defined SYS_PY (echo Python 3.9+ not found - install it from python.org & exit /b 1)
for %%f in (confirmations open_pos.csv vendor_master.csv beacon_erp.db) do (
  if not exist "%DATA%\%%f" (echo Missing %DATA%\%%f - usage: run_all.bat C:\path\to\data & exit /b 1)
)
where tesseract >nul 2>nul || echo WARNING: tesseract not on PATH - the 4 scanned PDFs will be listed as 'could not read'. See README.

echo == Installing Python packages into .venv (first run takes a minute)
if not exist "%HERE%.venv\Scripts\python.exe" %SYS_PY% -m venv "%HERE%.venv" || exit /b 1
set "PY=%HERE%.venv\Scripts\python.exe"
"%PY%" -m pip install -q --upgrade pip >nul
"%PY%" -m pip install -q -r "%HERE%task1_confirmation_recon\requirements.txt" || exit /b 1
if not exist "%OUT%" mkdir "%OUT%"

echo == Task 1 - PO confirmation check
cd /d "%HERE%task1_confirmation_recon"
"%PY%" recon.py --confirmations "%DATA%\confirmations" --pos "%DATA%\open_pos.csv" --vendors "%DATA%\vendor_master.csv" --erp "%DATA%\beacon_erp.db" --out "%OUT%" || exit /b 1
echo == Vendor forms
"%PY%" vendor_forms.py ack --pos "%DATA%\open_pos.csv" --vendors "%DATA%\vendor_master.csv" --erp "%DATA%\beacon_erp.db" --out "%OUT%\vendor_forms\acknowledgment_forms" >nul || exit /b 1
"%PY%" vendor_forms.py onboard --pos "%DATA%\open_pos.csv" --vendors "%DATA%\vendor_master.csv" --erp "%DATA%\beacon_erp.db" --out "%OUT%\vendor_forms\supplier_packs" >nul || exit /b 1
echo == Task 2 - vendor readout
cd /d "%HERE%task2_vendor_readout"
"%PY%" build_readout.py --erp "%DATA%\beacon_erp.db" --out "%OUT%" || exit /b 1
echo == Tests
set "BEACON_DATA=%DATA%"
cd /d "%HERE%task1_confirmation_recon"
"%PY%" -m unittest discover -s tests
cd /d "%HERE%task2_vendor_readout"
"%PY%" -m unittest discover -s tests
echo.
echo Done. Outputs are in %OUT%
endlocal
