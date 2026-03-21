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
echo === Step 1a: Installing pythonnet 2.5.2 (pre-built wheel, no compiler needed) ===
echo.
venv\Scripts\pip install "pythonnet==2.5.2"
if %errorlevel% neq 0 (
    echo ERROR: pythonnet install failed.
    pause
    exit /b 1
)

echo.
echo === Step 1b: Installing pyadomd (no-deps, uses the pythonnet above) ===
echo.
venv\Scripts\pip install pyadomd --no-deps
if %errorlevel% neq 0 (
    echo ERROR: pyadomd install failed.
    pause
    exit /b 1
)

echo.
echo === pyadomd + pythonnet installed OK ===

echo.
echo === Step 2: Running WK12 backfill ===
echo    (Power BI must be open with the FTT-DPV report refreshed)
echo.
venv\Scripts\python.exe update_morning_meeting.py --backfill-week

echo.
echo Done! Check mm_update_log.txt for details.
echo The Week Compare panel will now show all of last week.
pause
