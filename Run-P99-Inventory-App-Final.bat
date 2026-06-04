@echo off
cd /d "%~dp0"
python "%~dp0p99_inventory_app_final.py"
if errorlevel 1 py "%~dp0p99_inventory_app_final.py"
pause
