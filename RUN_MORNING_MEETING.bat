@echo off
:: ── A-Shop Morning Meeting Dashboard — Daily Auto-Update ──────────────────────
:: Runs update_morning_meeting.py using the project venv (has all packages)
:: Scheduled: Mon-Fri 07:45 via Windows Task Scheduler (AShop_MorningMeeting_Update)

cd /d "C:\Users\NYOUSIF\Desktop\AShop_Dashboard"

echo [%date% %time%] Starting morning meeting update... >> mm_update_log.txt

venv\Scripts\python.exe update_morning_meeting.py

if %errorlevel% == 0 (
    echo [%date% %time%] Morning meeting update completed successfully >> mm_update_log.txt
) else (
    echo [%date% %time%] Morning meeting update FAILED (exit code %errorlevel%) >> mm_update_log.txt
)
