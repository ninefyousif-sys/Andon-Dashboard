@echo off
:: ── STEP 1: GitHub Pages Setup ───────────────────────────────────────────────
:: Run this ONCE to connect the dashboard folder to GitHub Pages.
:: Before running: create an empty repo on github.com (e.g. "ashop-dashboard")
:: Then run this file as Administrator.

set REPO_URL=https://github.com/YOUR_USERNAME/ashop-dashboard.git
set FOLDER=C:\Users\NYOUSIF\Desktop\AShop_Dashboard

echo Initialising git in %FOLDER%...
cd /d %FOLDER%
git init
git add body_shop_intelligence.html
git commit -m "Initial dashboard publish"
git branch -M main
git remote add origin %REPO_URL%
git push -u origin main

echo.
echo === GitHub setup complete ===
echo 1. Go to your GitHub repo → Settings → Pages
echo 2. Set Source: Deploy from branch → main → / (root)
echo 3. Your dashboard will be live at:
echo    https://YOUR_USERNAME.github.io/ashop-dashboard/body_shop_intelligence.html
echo.
echo Now edit update_dashboard.py and set:
echo    GITHUB_ENABLED = True
echo.
pause
