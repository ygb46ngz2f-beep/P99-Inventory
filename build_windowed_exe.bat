@echo off
cd /d "%~dp0"
py -m pip install pyinstaller openpyxl
py -m PyInstaller --onefile --windowed --noupx p99_inventory_app_final.py
pause
