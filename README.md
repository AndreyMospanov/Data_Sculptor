# Data Sculptor

CLI backend for batch file renaming by masks and XLSX to CSV conversion.

## Install

Install optional XLSX editor libraries manually:

```powershell
python -m pip install -r requirements.txt
```

If `openpyxl` is missing, the GUI still starts, but the `XLSX editor` tab will ask you to run this command manually.

## Desktop GUI

Run the desktop application:

```powershell
python -m core.gui
```

On Windows, you can also start [data_sculptor_gui.pyw](data_sculptor_gui.pyw) by double-clicking it.
If double-clicking `.pyw` does not open anything, run [start_data_sculptor_gui.bat](start_data_sculptor_gui.bat) from the same folder to see startup errors in a console.

The GUI has three tabs:

- `Batch rename` previews and applies mask-based renames.
- `XLSX to CSV` converts one workbook or a folder with many `.xlsx` files. It supports comma, semicolon, tab, or a custom one-character delimiter.
- `XLSX editor` quickly checks header columns, inserts columns, deletes columns, and swaps two columns in one workbook or many workbooks in a folder. Insert, delete, and swap controls are separated into their own blocks.
- `CSV editor` provides the same header inspection, insert, delete, and swap operations for one CSV file or a folder of CSV files.

Use the `Language` menu to switch the main interface between English and Russian without restarting the app.

During XLSX conversion the GUI shows the current file, progress for the current file, and a `Cancel` button.

See [ARCHITECTURE.md](ARCHITECTURE.md) for a detailed architecture overview.

The XLSX editor shows progress while checking, inserting, or deleting columns. After `Check columns`, rows are selected by default. Use `Select all files` to select or clear every row, or click any row to toggle `[x]` / `[ ]`; insert and delete operations apply only to selected rows. It uses `openpyxl` for safer workbook edits and changes selected workbooks in place, so make a backup before mass editing important files.

CSV editor uses the standard `csv` module and rewrites files through a temporary file before replacing the original. The default `utf-8-sig` encoding writes a BOM so CSV files open correctly in Excel with Cyrillic text.

## Batch Renaming

Preview only:

```powershell
python -m core.cli "C:\path\to\files" --input "{name} {date:DD.MM.YYYY}.{ext}" --output "{date:YYYYMMDD}_{name}.{ext}"
```

Apply changes:

```powershell
python -m core.cli "C:\path\to\files" --input "{name} {date:DD.MM.YYYY}.{ext}" --output "{date:YYYYMMDD}_{name}.{ext}" --apply
```

## Masks

Placeholders are written as `{field}` or `{field:FORMAT}`.

Example input:

```text
{name} {date:DD.MM.YYYY}.{ext}
```

Example output:

```text
{date:YYYYMMDD}_{name}.{ext}
```

This converts `Report 31.05.2026.txt` to `20260531_Report.txt`.

Supported date tokens: `DD`, `MM`, `YY`, `YYYY`.

Useful options:

- `--apply` renames files; without it the command only previews.
- `--recursive` processes nested directories.
- `--include-dirs` includes directories.
- `--overwrite` allows replacing existing target files.

## XLSX to CSV

List workbook sheets:

```powershell
python -m core.xlsx_cli "C:\path\to\input.xlsx" --list-sheets
```

Convert the first sheet with comma delimiter:

```powershell
python -m core.xlsx_cli "C:\path\to\input.xlsx" "C:\path\to\output.csv"
```

Convert a selected sheet with semicolon delimiter:

```powershell
python -m core.xlsx_cli "C:\path\to\input.xlsx" "C:\path\to\output.csv" --sheet "Sheet1" --delimiter ";"
```

Convert every `.xlsx` file in a folder:

```powershell
python -m core.xlsx_cli "C:\path\to\xlsx-folder" "C:\path\to\csv-folder" --delimiter ";"
```

Convert nested folders too:

```powershell
python -m core.xlsx_cli "C:\path\to\xlsx-folder" "C:\path\to\csv-folder" --delimiter ";" --recursive
```

Tab delimiter can be passed as `--delimiter "\t"`.

Excel dates are detected automatically when the XLSX cell has a date style. For example, Excel serial value `44953` is exported as `27.01.2023` instead of a raw number.
