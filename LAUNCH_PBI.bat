@echo off
REM ═══════════════════════════════════════════════════════════════════
REM  LAUNCH_PBI.bat
REM  Opens both Power BI files needed for the morning meeting update.
REM  Scheduled: Mon-Fri 7:20am  (so PBI is loaded before 7:45am run)
REM ═══════════════════════════════════════════════════════════════════

REM ── A Shop FTT / DPV / W&G ────────────────────────────────────────
start "" "C:\Users\NYOUSIF\OneDrive - Volvo Cars\Desktop\Power BI\FTT-DPV (5).pbix"

REM ── BIW SQD Axxos (OPR / Scrap) ──────────────────────────────────
start "" "C:\Users\NYOUSIF\OneDrive - Volvo Cars\Desktop\Power BI\BIW SQD Axxos.pbix"

timeout /t 3 /nobreak >nul

echo [%date% %time%] PBI files launched >> "C:\Users\NYOUSIF\Desktop\AShop_Dashboard\mm_update_log.txt"
