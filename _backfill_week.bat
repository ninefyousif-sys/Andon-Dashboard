@echo off
REM ═══════════════════════════════════════════════════════════════════
REM  BACKFILL WEEK (D1-D5)
REM  Re-runs all working days Mon-Fri through the fixed script
REM  and repopulates the dashboard history.
REM
REM  REQUIREMENTS:
REM    1. Both Power BI files must be OPEN:
REM         - A Shop Body Count (port 58016)
REM         - BIW SQD Axxos    (port 59110)
REM    2. Run this from the AShop_Dashboard folder
REM ═══════════════════════════════════════════════════════════════════

cd /d "C:\Users\NYOUSIF\Desktop\AShop_Dashboard"

echo.
echo ============================================================
echo  BACKFILL WK12 (Mon 16-Mar through Fri 20-Mar)
echo  Make sure BOTH Power BI files are open first!
echo ============================================================
echo.
pause

python update_morning_meeting.py --backfill-week

echo.
echo ============================================================
echo  Done! Check mm_update_log.txt for results.
echo ============================================================
pause
