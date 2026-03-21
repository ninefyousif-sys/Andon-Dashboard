@echo off
:: ══════════════════════════════════════════════════════════════════
::  Discover BOK/BOL OPR table names from BIW SQD Axxos
::
::  Before running:
::   - Open BIW SQD Axxos in Power BI Desktop and refresh it
::   - FTT-DPV can also be open — both will be scanned
::
::  After running:
::   - Open mm_update_log.txt
::   - Find the TABLE: and COLUMN: lines for BOK/BOL/OPR tables
::   - Fill them into update_morning_meeting.py at the top:
::       OPR_TABLE    = 'YOUR TABLE NAME'
::       OPR_COL_BOK  = 'YOUR BOK COLUMN'
::       OPR_COL_BOL  = 'YOUR BOL COLUMN'
:: ══════════════════════════════════════════════════════════════════

cd /d "C:\Users\NYOUSIF\Desktop\AShop_Dashboard"

echo.
echo === Discovering PBI table names for BOK/BOL OPR ===
echo    Make sure BIW SQD Axxos is open and refreshed in Power BI Desktop.
echo.
venv\Scripts\python.exe update_morning_meeting.py --discover-tables

echo.
echo Done! Open mm_update_log.txt to see all TABLE: and COLUMN: entries.
echo Then fill in OPR_TABLE / OPR_COL_BOK / OPR_COL_BOL in update_morning_meeting.py
pause
