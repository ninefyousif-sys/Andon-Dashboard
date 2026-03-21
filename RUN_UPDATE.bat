@echo off
:: Body Shop Dashboard — Daily Update Launcher
:: Uses project venv (venv\Scripts\python.exe) — always has correct packages

cd /d "C:\Users\NYOUSIF\Desktop\AShop_Dashboard"

echo [%date% %time%] Starting dashboard update... >> update_log.txt

:: Use the project virtualenv — no PATH/venv conflicts
venv\Scripts\python.exe update_dashboard.py >> update_log.txt 2>&1

if %errorlevel% == 0 (
    echo [%date% %time%] Update completed successfully >> update_log.txt
) else (
    echo [%date% %time%] Update FAILED (exit code %errorlevel%) >> update_log.txt
)
