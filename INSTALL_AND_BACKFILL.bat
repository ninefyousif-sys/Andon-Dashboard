@echo off
:: ══════════════════════════════════════════════════════════════════
::  A-Shop Dashboard — Backfill + BOK/BOL OPR Setup
::  No extra pip install needed — uses PowerShell COM/ADODB.
::
::  STEP 1: Run WK backfill (FTT, DPV, downtime, PPT)
::     Open Power BI Desktop with FTT-DPV file refreshed.
::     Then run this bat or run step 1 below.
::
::  STEP 2: Find BOK/BOL OPR table names
::     Open BIW SQD Axxos in Power BI Desktop as well.
::     Then run DISCOVER_OPR_TABLES.bat (or the command below).
::     Paste the table names into update_morning_meeting.py:
::       OPR_TABLE    = 'YOUR TABLE NAME'
::       OPR_COL_BOK  = 'YOUR BOK COLUMN'
::       OPR_COL_BOL  = 'YOUR BOL COLUMN'
::     Then re-run this bat to backfill with OPR values.
:: ══════════════════════════════════════════════════════════════════

cd /d "C:\Users\NYOUSIF\Desktop\AShop_Dashboard"

echo.
echo === Running WK backfill (FTT / DPV / Downtime / PPT) ===
echo    Power BI Desktop must be open with the report refreshed.
echo.
venv\Scripts\python.exe update_morning_meeting.py --backfill-week

echo.
echo Done! Check mm_update_log.txt for details.
echo.
echo ─────────────────────────────────────────────────────────
echo NEXT STEP — Find BOK/BOL OPR table names:
echo   1. Open BIW SQD Axxos in Power BI Desktop (refresh it)
echo   2. Run: DISCOVER_OPR_TABLES.bat
echo   3. Fill OPR_TABLE / OPR_COL_BOK / OPR_COL_BOL into
echo      update_morning_meeting.py (top of file)
echo   4. Run this bat again to backfill with OPR data
echo ─────────────────────────────────────────────────────────
pause
