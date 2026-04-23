@echo off
chcp 65001 >nul
cd /d "C:\Users\user\Desktop\順丰E順递"

echo [%DATE% %TIME%] 開始定時更新 >> logs\scheduled_update.log

python scripts\scheduled_update.py >> logs\scheduled_update.log 2>&1

echo [%DATE% %TIME%] 完成 >> logs\scheduled_update.log
