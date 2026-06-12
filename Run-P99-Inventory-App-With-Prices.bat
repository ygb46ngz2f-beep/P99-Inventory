@echo off
cd /d "%~dp0"
python "%~dp0p99_inventory_app_final_with_prices.py"
if errorlevel 1 py "%~dp0p99_inventory_app_final_with_prices.py"
pause
