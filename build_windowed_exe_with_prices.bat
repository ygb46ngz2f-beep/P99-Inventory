@echo off
cd /d "%~dp0"
py -m pip install pyinstaller openpyxl requests
py -m PyInstaller --onefile --windowed --noupx p99_inventory_app_final_with_prices.py
pause
