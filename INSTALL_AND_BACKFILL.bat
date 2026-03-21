@echo off
:: ══════════════════════════════════════════════════════════════════
::  Run WK12 backfill — no pip install needed.
::  DAX queries now use PowerShell COM/ADODB (built into Windows).
::
::  Requires: Power BI Desktop open with FTT-DPV or BIW SQD Axxos
:: ══════════════════════════════════════════════════════════════════

cd /d "C:\Users\NYOUSIF\Desktop\AShop_Dashboard"

echo.
echo === Running WK12 backfill ===
echo    Power BI must be open with the report refreshed.
echo.
venv\Scripts\python.exe update_morning_meeting.py --backfill-week

echo.
echo Done! Check mm_update_log.txt for details.
echo The Week Compare panel will now show all of last week.
pause
