Files included:
- p99_inventory_app_final_with_prices.py : final Python app with optional price-data URL
- Run-P99-Inventory-App-With-Prices.bat : runs the Python app directly
- build_console_exe_with_prices.bat : builds a console EXE for debugging
- build_windowed_exe_with_prices.bat : builds a normal windowed EXE
- README-P99-With-Prices-Full.txt : this file

Recommended order:
1. Run build_console_exe_with_prices.bat
2. Test the EXE from the dist folder in PowerShell with:
   .\p99_inventory_app_final_with_prices.exe
3. If it works, run build_windowed_exe_with_prices.bat

Price-data URL:
- The app has a field where you can paste a JSON URL.
- Default: https://www.pigparse.org/api/item/getall/Green
- The app waits 20 seconds before downloading the JSON.

Required on Windows:
- Python installed
- py launcher available
- Internet access for pip install and price-data download
