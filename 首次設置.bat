@echo off
chcp 65001 > nul
echo ========================================
echo 順丰寄件自動化 - 首次設置
echo ========================================
cd /d "%~dp0scripts"
python setup.py
pause
