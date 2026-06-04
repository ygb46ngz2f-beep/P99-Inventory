# P99 Inventory to Excel

P99 Inventory to Excel is a beginner-friendly desktop tool for Project 1999 players that turns inventory `.txt` files into a clean Excel workbook. It creates one tab per file, builds an item summary across all imported files, adds wiki links, and highlights items that match your chosen keywords.

This project is made for players who want an easier way to organize bag, inventory, and bank data without manually copying everything into Excel.

## What it does

- Imports one or more Project 1999 inventory `.txt` files
- Creates one worksheet per imported file
- Updates existing sheets instead of making duplicate tabs
- Builds an **Item Summary** sheet across all imported files
- Adds Project 1999 wiki links for item names
- Highlights full rows when the item name matches selected keywords
- Lets you add your own custom highlight keywords
- Sorts highlighted items to the top of the summary sheet

## Who this is for

This tool is for players who:
- want to compare multiple characters or inventories
- want to track duplicate items across files
- want important items highlighted automatically
- do not want to manually format raw inventory text in Excel

## Files included

You may see some or all of these files in the project:

- `p99_inventory_app_final.py` — the main Python app
- `Run-P99-Inventory-App-Final.bat` — runs the Python app directly
- `build_console_exe.bat` — builds a debug `.exe` that shows errors
- `build_windowed_exe.bat` — builds the normal double-click `.exe`
- `README.md` — this guide

## Before you start

You need:

1. A Windows computer
2. Python installed
3. Excel or another program that can open `.xlsx` files
4. Your Project 1999 inventory `.txt` files

To check whether Python is installed:

```bat
py --version
```

If that does not work, install Python from the official site and make sure the `py` launcher is available.

## Quick start

If you just want to run the Python version:

1. Download the project files into one folder.
2. Put your inventory `.txt` files wherever you like.
3. Double-click `Run-P99-Inventory-App-Final.bat`
4. Add your `.txt` files in the app
5. Choose where to save the Excel workbook
6. Click **Create / Update Workbook**

## First-time setup

Open **PowerShell** or **Command Prompt** in the project folder and install the required package:

```bat
py -m pip install openpyxl
```

If you want to build the `.exe` version too, install PyInstaller:

```bat
py -m pip install pyinstaller openpyxl
```

## How to run the app

### Option 1: Run the Python version

Double-click:

```txt
Run-P99-Inventory-App-Final.bat
```

This is the easiest way to test the app before building an `.exe`.

### Option 2: Run from terminal

Open PowerShell in the project folder and run:

```bat
py p99_inventory_app_final.py
```

## How to use the app

1. Open the app
2. Click **Add TXT Files**
3. Select one or more inventory `.txt` files
4. Choose the Excel output file location
5. Add or remove highlight keywords if needed
6. Click **Create / Update Workbook**

The app will then:
- import each file into its own sheet
- replace an existing sheet if the same file was imported before
- rebuild the summary sheet
- highlight matching rows
- save the Excel workbook

## How the Excel file is organized

### Character sheets
Each imported file gets its own worksheet with:
- Location
- Name
- ID
- Count
- Slots
- Wiki Link

### Item Summary
The summary sheet combines items from all imported sheets and shows:
- item name
- item ID
- total count across files
- which sheets the item was found in
- wiki link

Highlighted summary rows are sorted to the top.

## Highlight keywords

The app starts with these default keywords:

- Crown
- Throne
- Squire
- Knight

You can also:
- add your own keywords
- remove keywords you do not want
- use partial matches inside longer item names

If a keyword appears anywhere in the item name, the whole row is highlighted.

## Updating an existing workbook

You can reuse the same Excel file over and over.

When you import a file again:
- if the sheet name already exists, the app replaces that sheet
- the summary sheet is rebuilt using the latest imported data
- highlighted items are recalculated

This means you do not need to start from scratch every time.

## How to build the EXE

### Step 1: Build a debug EXE

Start with the console build so you can see any error messages:

```txt
build_console_exe.bat
```

Or run manually:

```bat
py -m PyInstaller --onefile --console --noupx p99_inventory_app_final.py
```

### Step 2: Test the EXE

Go into the `dist` folder and run:

```powershell
.\p99_inventory_app_final.exe
```

If it works, move to the next step.

### Step 3: Build the normal EXE

Run:

```txt
build_windowed_exe.bat
```

Or manually:

```bat
py -m PyInstaller --onefile --windowed --noupx p99_inventory_app_final.py
```

That creates a normal double-clickable GUI version.

## Troubleshooting

### The app says `ModuleNotFoundError`
Install the required package:

```bat
py -m pip install openpyxl
```

### The EXE does not open when I double-click it
Build the console version first and run it from PowerShell:

```powershell
.\p99_inventory_app_final.exe
```

That usually shows the real error.

### PowerShell says the EXE is not recognized
Use `.\` before the filename:

```powershell
.\p99_inventory_app_final.exe
```

### I get `Table with name Table_X already exists`
Use the newest version of the script. The final version creates unique workbook-wide table names and replaces existing sheets instead of stacking old table objects.

### My old sheet was not updated
Make sure the file has the same base filename as before. The worksheet name is based on the imported file’s name.

## Example workflow

Example:
1. Export inventory text from Character A
2. Export inventory text from Character B
3. Open the app
4. Add both `.txt` files
5. Save to `p99_inventory.xlsx`
6. Review each character tab
7. Open the summary sheet to see shared items and totals

## Why this tool exists

Project 1999 inventory exports are useful, but raw text is hard to read and compare. This app makes those exports easier to browse, search, and track in Excel.

## Notes

- Designed for Windows
- Built with Python, Tkinter, and openpyxl
- Excel is recommended for the best table and formatting support

## License

Add your preferred license here, such as MIT.

## Contact

Add your GitHub name or preferred contact info here.
