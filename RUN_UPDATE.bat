@echo off
:: Body Shop Dashboard — Daily Update Launcher

cd /d "C:\Users\NYOUSIF\Desktop\AShop_Dashboard"

echo [%date% %time%] Starting dashboard update... >> update_log.txt

:: Force Python user site-packages
set PYTHONPATH=C:\Users\NYOUSIF\AppData\Roaming\Python\Python314\site-packages

:: Run with Python 3.14 — quoted to handle "Program Files" space
"%ProgramFiles%\Python314\python.exe" update_dashboard.py >> update_log.txt 2>&1

if %errorlevel% == 0 (
    echo [%date% %time%] Update completed successfully >> update_log.txt
) else (
    echo [%date% %time%] Update FAILED (exit code %errorlevel%) >> update_log.txt
)
