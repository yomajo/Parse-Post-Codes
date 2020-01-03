# Parse Post Codes


## Description & Usage

This a part of a large VBA Project I was working on, that generates Invoices and Export Declarations from exported Ebay CSV Files. Each Export Declaration uses barcode, that is a special font used on postal code. Postal Codes are provided by Post Office in txt format.

Example of first 4 codes in text file:

"UA759883946LTUA745082482LTUA860409989LTUA432531574LT"

*Postal code structure: 'UA' + 9 dec symbols + 'LT'*

### Usage

`Main Program.xlsm` works as GUI. Screenshot from "Main" sheet in `Main Program.xlsm`:
![Parse Post Codes Excel GUI screenshot](https://user-images.githubusercontent.com/45366313/71713394-2b5c9c00-2e12-11ea-8122-3f90849b8275.JPG)


1. On workbook open, command prompt is instantiated and `where python` is executed, capturing output to hidden sheet "PyPath" in `A1` cell.
2. **Reset** button simply cleans contents from *Main* worksheet
3. **Use New Code** `Postal_Codes_Manager.xlsx`, copies first free code and moves it to "Expired Codes" sheet
4. **Add New Postal Codes** launches `txt_to_excel.py`, which converts txt file contents to list and starts appending it from first free row in A column, "Free Codes" sheet, `Postal_Codes_Manager.xlsx` workbook

### Python to parse codes and load to workbook

Script splits txt continuous content to list each 13 symbols (fixed) and via openpyxl lib. pushes new codes in `Postal_Codes_Manager.xlsx`

### Python to generate postal codes

Not to bother Post Office when still in development, I postal codes were still needed to work with, so `output_entropy.py` was written to generate identical file of post codes. It generates txt file with `codes_count` (Global variable inside script) number of postal codes.

Example [file](post_codes.txt).

### VBA Modules

All VBA part modules have been exported for inspection and are in [VBA Modules](Parse-Post-Codes/tree/master/VBA%20modules)


## Requirements

- Python 3.7.3+
- openpyxl (`pip install openpyxl`)
- Python added to PATH
- In order to use barcode, add provided [font](IDAutomationHC39M%20Free%20Version.otf) to C:\Windows\Fonts.


*Tested only on Windows machine*

## Installation

Simply download both workbooks and python files and proceed with `Main Program.xlsm` as GUI.
