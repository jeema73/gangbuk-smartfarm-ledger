@echo off
cd /d "%~dp0"
chcp 65001 > nul
echo ====================================================
echo  [SmartFarm] Pushing code to GitHub...
echo ====================================================
echo.
git branch -M main
git push -u origin main
echo.
echo ====================================================
pause
