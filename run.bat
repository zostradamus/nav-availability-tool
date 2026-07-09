@echo off
cd /d "%~dp0"
python run_nav_analysis.py >> nav_log.txt 2>&1
