@echo off
:: ══════════════════════════════════════════════════════════════════
::  STEP 1 — Install pyadomd (needed to query Power BI Desktop)
::  STEP 2 — Run WK12 backfill so Week Compare panel shows last week
::
::  Run this ONCE from the AShop_Dashboard folder.
::  Requires: Power BI Desktop open with FTT-DPV (or BIW SQD Axxos)
:: ══════════════════════════════════════════════════════════════════

cd /d "C:\Users\NYOUSIF\Desktop\AShop_Dashboard"

echo.
echo === Step 1: Installing pyadomd into project venv ===
echo.
venv\Scripts\pip install pyadomd
if %errorlevel% neq 0 (
    echo ERROR: pip install failed. Check your internet connection.
    pause
    exit /b 1
)
echo pyadomd installed OK.

echo.
echo === Step 2: Running WK12 backfill ===
echo    (Power BI must be open with the FTT-DPV report refreshed)
echo.
venv\Scripts\python.exe update_morning_meeting.py --backfill-week

echo.
echo Done! Check mm_update_log.txt for details.
echo The Week Compare panel will now show all of last week.
pause
