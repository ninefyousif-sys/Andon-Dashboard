@echo off
:: Body Shop Dashboard — Daily Update Launcher
:: Called by Windows Task Scheduler at 17:00 Mon-Fri
:: Reads HOP + DT Excel files, updates dashboard HTML, pushes to GitHub

cd /d "C:\Users\NYOUSIF\Desktop\AShop_Dashboard"

echo [%date% %time%] Starting dashboard update... >> update_log.txt

:: Run the Python update script
python update_dashboard.py >> update_log.txt 2>&1

if %errorlevel% == 0 (
    echo [%date% %time%] Update completed successfully >> update_log.txt
) else (
    echo [%date% %time%] Update FAILED (exit code %errorlevel%) >> update_log.txt
)
