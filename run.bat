@echo off
cd /d "%~dp0"
python -u run_nav_analysis.py >> nav_log.txt 2>&1
