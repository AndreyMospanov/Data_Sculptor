# Data Sculptor
<div align="center">
  <p>
    <a href="#"><img src="./img/icon.ico" height="150px" alt="logo" /></a>
  </p>
</div>

**Read this in other languages / Читайте на других языках**
- **[English](README.md)**
- [Русский](README.ru.md)

Convert unstructured source data in files with one click
[v] Mass renaming of sample files using dates
[v] Mass conversion of xlsx->csv files with choice of encoding and separator
[v] Mass checking of csv, xlsx file structure
[v] Massive change in the structure of xlsx and csv files with the ability to select specific files from the list

## Install

Download the latest release for Windows from the [releases page](https://github.com/AndreyMospanov/Data_Sculptor/releases).

To use Python:
Install libraries:

```powershell
python -m pip install -r requirements.txt
```
Start the GUI:

```powershell
python -m core.gui
```
or double click [data_sculptor_gui.pyw](data_sculptor_gui.pyw) to start the GUI and start work.
If double-clicking `.pyw` does not open anything, run [start_data_sculptor_gui.bat](start_data_sculptor_gui.bat) from the same folder to see startup errors in a console.
Optionally be sure that You've installed the required libraries.

## License

This project is licensed under the MIT License - see the 
[LICENSE](https://opensource.org/licenses/MIT) for details.

## Desktop GUI

The GUI has some tabs:

- `Batch rename` previews and applies mask-based renames.
- `XLSX to CSV` converts one workbook or a folder with many `.xlsx` files. It supports comma, semicolon, tab, or a custom one-character delimiter.
- `XLSX editor` quickly checks header columns, inserts columns, deletes columns, and swaps two columns in one workbook or many workbooks in a folder. Insert, delete, and swap controls are separated into their own blocks.
- `CSV editor` provides the same header inspection, insert, delete, and swap operations for one CSV file or a folder of CSV files.

Use the `Language` menu to switch the main interface between English and Russian without restarting the app.

During XLSX conversion the GUI shows the current file, progress for the current file, and a `Cancel` button.

See [ARCHITECTURE.md](ARCHITECTURE.md) for a detailed architecture overview.

The Application shows progress while checking, inserting, or deleting columns. It also has a status bar for all tabs, be sure to check if you think, that app is not responding or freeze.
