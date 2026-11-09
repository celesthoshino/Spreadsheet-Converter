# Spreadsheet Converter

Converts spreadsheets and data files between Excel, CSV, TSV, OpenDocument, JSON and HTML. Put your files in the
`input` folder, run the script, and get the converted files in the `output` folder.

It's careful with your data. ZIP codes and phone numbers keep their leading zeros, long ID numbers aren't rounded,
accents survive, and Excel workbooks with several sheets are handled properly.

## Folder layout

```
spreadsheet_converter/
├── convert_sheets.py   the script
├── requirements.txt    Python packages to install
├── README.md           this file
├── input/              put the files you want converted here
└── output/             converted files appear here, one subfolder per format
    ├── xlsx/
    ├── csv/
    └── ...
```

## Setup (one time)

1. Install **Python 3.9 or newer** from https://www.python.org/downloads/
   (on Windows, tick "Add Python to PATH" during install).
2. Open a terminal in the `spreadsheet_converter` folder and run:

   ```
   pip install -r requirements.txt
   ```

You don't need Excel installed.

## How to use

1. Copy the files you want converted into the `input` folder.
2. Open a terminal in the `spreadsheet_converter` folder and run:

   ```
   python convert_sheets.py
   ```

3. It lists the files it found and asks which formats you want. Type one or more, separated by spaces (for example
   `xlsx csv`), or type `all`.
4. The converted files are saved in `output/<format>/`. For example, `input/report.csv` converted to Excel becomes
   `output/xlsx/report.xlsx`.

| Command | What it does |
|---|---|
| `python convert_sheets.py` | Asks which formats you want |
| `python convert_sheets.py xlsx` | Converts everything to Excel |
| `python convert_sheets.py csv json` | Converts everything to CSV and JSON |
| `python convert_sheets.py --all` | Converts everything to every format |
| `python convert_sheets.py --help` | Shows all options |

## Options

| Option | Default | Meaning |
|---|---|---|
| `formats` | asks you | One or more formats to convert to |
| `--all` | off | Convert to every supported format |
| `-i`, `--input` | `input` folder | Folder to read files from |
| `-o`, `--output` | `output` folder | Folder to save converted files to |

## Supported formats

| Format | Can read | Can write | Notes |
|---|---|---|---|
| `xlsx` | Yes (also `.xlsm` and old `.xls`) | Yes | Excel. Columns are sized to fit and the heading row stays visible when scrolling |
| `ods` | Yes | Yes | OpenDocument, used by LibreOffice and OpenOffice |
| `csv` | Yes | Yes | Plain text, one row per line. Opens in any spreadsheet program |
| `tsv` | Yes (also `.tab`) | Yes | Like CSV, but separated by tabs |
| `json` | Yes | Yes | Used by programs and websites. Each row becomes `{"column": value, ...}` |
| `html` | No | Yes | A web page with the table, to view in a browser or paste into an email |

Other files in `input` are listed as "ignored" and left alone.

## Workbooks with several sheets

- **To xlsx, ods, json or html:** all sheets are kept in one file.
- **To csv or tsv:** these formats only hold one table, so each sheet becomes its own file:
  `budget.xlsx` with sheets "2023" and "2024" becomes `budget_2023.csv` and `budget_2024.csv`.

## How your data is kept safe

- **Leading zeros are kept.** ZIP codes like `01234` and phone numbers like `0412 345 678` stay exactly as they are.
- **Long numbers aren't rounded.** Numbers longer than 15 digits (credit card numbers, some IDs) are kept as text,
  because Excel would otherwise round them and change the last digits.
- **Only clearly numeric columns become numbers** when reading CSV files. Anything unclear stays as text, exactly as
  written.
- **Accents and symbols survive.** CSV files from older Excel versions (which don't use UTF-8) are read correctly, and
  the CSV files this script writes open correctly in Excel.
- **Semicolon CSVs work.** Excel in many European countries saves CSV with `;` instead of `,`, and this is detected
  automatically.
- **Dates** are written as `2024-05-14` (or `2024-05-14 13:45:00` if they have a time) in text formats.

## Good to know

- **Original files are never changed or deleted.** The script only reads from `input`.
- **Only the data is converted, not the formatting.** Colors, fonts, merged cells, charts and formulas are not kept.
  For formulas, the last calculated value is used.
- **The first row is used as the column headings.** Columns with no heading get the names `column_1`, `column_2`, ...
- **Dates from CSV and JSON files stay as text** when converted to Excel, since those formats don't mark which values
  are dates.
- **JSON with nested data** like `{"address": {"city": "Oslo"}}` becomes a column named `address.city`.
- **Converting a file to the format it's already in is skipped.**
- **Running it again overwrites earlier results** with the same name in `output`.
- **Files with the same name but different extensions**, like `data.csv` and `data.xlsx`, keep their old extension in
  the new name so neither overwrites the other: `data_csv.json` and `data_xlsx.json`.
- **Subfolders inside `input` are not scanned.** Put the files directly in `input`.

## Troubleshooting

**`Error: pandas is not installed` or `a package is missing`**

Run `pip install -r requirements.txt` from inside the `spreadsheet_converter` folder.

**`it could not be read`**

The file is damaged, password protected, or not really a spreadsheet (for example, a web page saved with a `.xls`
name). Try opening it in Excel and saving it again as `.xlsx`.

**A CSV opens in Excel with everything in one column**

Your Excel expects `;` between values. Convert to `xlsx` instead. Excel files open correctly everywhere.

**`'python' is not recognized`**

Python isn't on your PATH. On Windows, try `py convert_sheets.py` instead, or reinstall Python with "Add Python to PATH" ticked.
