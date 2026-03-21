@echo off
REM ══════════════════════════════════════════════════════════════════
REM  BACKFILL WK12 HISTORY — Run once to populate the Week Compare
REM  panel with last week's data (Mon–Fri).
REM
REM  BEFORE running:
REM    1. Open Power BI Desktop
REM    2. Open the FTT-DPV report and click Refresh
REM    3. Then double-click this file
REM ══════════════════════════════════════════════════════════════════

cd /d "C:\Users\NYOUSIF\Desktop\AShop_Dashboard"
echo.
echo === Backfilling WK12 history using Power BI + OneDrive data ===
echo.

venv\Scripts\python.exe update_morning_meeting.py --backfill-week >> mm_update_log.txt 2>&1

echo.
echo Done! Check mm_update_log.txt for details.
echo The Week Compare panel will now show all of last week.
pause
