@echo off
cd /d "%~dp0"
py -m pip install pyinstaller openpyxl requests
py -m PyInstaller --onefile --console --noupx p99_inventory_app_wts_30d_only.py
pause
