@echo off
rem One command for graders on Windows:  run_all.bat C:\path\to\data
rem Installs into a local .venv, runs both tasks, the vendor forms and the tests. Outputs go to .\output
setlocal
set "DATA=%~1"
if "%DATA%"=="" set "DATA=C:\data"
set "HERE=%~dp0"
set "OUT=%HERE%output"

where python >nul 2>nul || (echo Python 3.9+ not found & exit /b 1)
for %%f in (confirmations open_pos.csv vendor_master.csv beacon_erp.db) do (
  if not exist "%DATA%\%%f" (echo Missing %DATA%\%%f - usage: run_all.bat C:\path\to\data & exit /b 1)
)
where tesseract >nul 2>nul || echo WARNING: tesseract not installed - the 4 scanned PDFs will be listed as 'could not read'.

if not exist "%HERE%.venv\Scripts\python.exe" python -m venv "%HERE%.venv"
set "PY=%HERE%.venv\Scripts\python.exe"
"%PY%" -m pip install -q -r "%HERE%task1_confirmation_recon\requirements.txt" || exit /b 1
if not exist "%OUT%" mkdir "%OUT%"

cd /d "%HERE%task1_confirmation_recon"
echo == Task 1 - PO confirmation check
"%PY%" recon.py --confirmations "%DATA%\confirmations" --pos "%DATA%\open_pos.csv" --vendors "%DATA%\vendor_master.csv" --erp "%DATA%\beacon_erp.db" --out "%OUT%"
echo == Vendor forms
"%PY%" vendor_forms.py ack --pos "%DATA%\open_pos.csv" --vendors "%DATA%\vendor_master.csv" --erp "%DATA%\beacon_erp.db" --out "%OUT%\vendor_forms\acknowledgment_forms" >nul
"%PY%" vendor_forms.py onboard --pos "%DATA%\open_pos.csv" --vendors "%DATA%\vendor_master.csv" --erp "%DATA%\beacon_erp.db" --out "%OUT%\vendor_forms\supplier_packs" >nul
cd /d "%HERE%task2_vendor_readout"
echo == Task 2 - vendor readout
"%PY%" build_readout.py --erp "%DATA%\beacon_erp.db" --out "%OUT%"
echo == Tests
set "BEACON_DATA=%DATA%"
cd /d "%HERE%task1_confirmation_recon" && "%PY%" -m unittest discover -s tests
cd /d "%HERE%task2_vendor_readout" && "%PY%" -m unittest discover -s tests
echo Done. Outputs are in %OUT%
