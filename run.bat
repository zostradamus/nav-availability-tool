@echo off
cd /d "%~dp0"
powershell -Command "python -u run_nav_analysis.py 2>&1 | ForEach-Object { Write-Host $_; Add-Content nav_log.txt $_ -Encoding UTF8 }"
