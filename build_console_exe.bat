@echo off
cd /d "%~dp0"
py -m pip install pyinstaller openpyxl
py -m PyInstaller --onefile --console --noupx p99_inventory_app_final.py
pause
