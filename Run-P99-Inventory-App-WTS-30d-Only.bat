@echo off
cd /d "%~dp0"
python "%~dp0p99_inventory_app_wts_30d_only.py"
if errorlevel 1 py "%~dp0p99_inventory_app_wts_30d_only.py"
pause
