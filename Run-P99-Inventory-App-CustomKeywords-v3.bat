@echo off
cd /d "%~dp0"
python "%~dp0p99_inventory_app_custom_keywords_v3.py"
if errorlevel 1 py "%~dp0p99_inventory_app_custom_keywords_v3.py"
pause
