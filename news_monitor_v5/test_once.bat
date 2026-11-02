@echo off
cd /d "%~dp0"
python news_monitor.py --once
pause
