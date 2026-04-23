@echo off
chcp 65001 >nul
cd /d "C:\Users\user\Desktop\順丰E順递"
git add data/tracking.xlsx scripts/
git commit -m "update %date% %time%"
git push origin main
echo.
echo ✅ 已同步到 GitHub
