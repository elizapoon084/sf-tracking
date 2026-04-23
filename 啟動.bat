@echo off
chcp 65001 > nul
cd /d "%~dp0scripts"
python main_gui.py
pause
