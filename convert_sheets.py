"""
Spreadsheet Converter
---------------------
Put spreadsheets in the "input" folder, run this script, and converted copies
show up in the "output" folder (one subfolder per format).

    python convert_sheets.py              # asks which formats you want
    python convert_sheets.py xlsx         # everything to Excel
    python convert_sheets.py csv json     # everything to CSV and JSON
    python convert_sheets.py --all        # everything to every format

See README.md for all the details.
"""

import argparse
import csv
import html
import json
import re
import sys
from collections import Counter
from datetime import date, datetime
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    sys.exit("Error: pandas is not installed.\nFix it by running:  pip install -r requirements.txt")

# The input/output folders live next to this script, so it works no matter
# which folder you run it from.
BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"

FORMATS = {"csv": ".csv", "tsv": ".tsv", "xlsx": ".xlsx", "ods": ".ods", "json": ".json", "html": ".html"}

# Input file extension -> format name.
INPUT_EXTENSIONS = {".csv": "csv", ".tsv": "tsv", ".tab": "tsv", ".xlsx": "xlsx", ".xlsm": "xlsm",
                    ".xls": "xls", ".ods": "ods", ".json": "json"}

# A number that's safe to store as a number: no leading zeros (so ZIP codes like 01234
# stay text) and at most 15 digits (Excel rounds anything longer, e.g. long ID numbers).
SAFE_NUMBER = r"-?(?:0|[1-9]\d{0,14})(?:\.\d+)?"

HTML_PAGE = """<!doctype html>
<html><head><meta charset="utf-8"><title>{title}</title>
<style>
  body {{ font-family: Segoe UI, Arial, sans-serif; margin: 2em; color: #222; }}
  table {{ border-collapse: collapse; margin-bottom: 2em; }}
  th, td {{ border: 1px solid #ccc; padding: 4px 10px; text-align: left; }}
  th {{ background: #f0f0f0; }}
  tr:nth-child(even) td {{ background: #fafafa; }}
</style></head>
<body>
{body}
</body></html>
"""


class SheetProblem(Exception):
    """A file that can't be read, with a message that explains why."""


# ---------------------------------------------------------------- reading

def detect_delimiter(path: Path, encoding: str) -> str:
    """Comma, semicolon, tab or pipe, whichever the file uses. Excel in many European countries uses semicolons."""
    with open(path, encoding=encoding, newline="") as f:
        sample = f.read(20000)
    try:
        return csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
    except csv.Error:
        return ","


def read_text_table(path: Path, delimiter=None) -> pd.DataFrame:
    """Read a CSV/TSV file. Tries UTF-8 first, then the Windows encoding older Excel versions use."""
    for encoding in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            sep = delimiter or detect_delimiter(path, encoding)
            # Everything is read as text first, so nothing gets changed by accident.
            table = pd.read_csv(path, sep=sep, dtype=str, keep_default_na=False, encoding=encoding)
            return text_to_numbers(table)
        except UnicodeDecodeError:
            continue
        except pd.errors.EmptyDataError:
            return pd.DataFrame()
    raise SheetProblem("couldn't work out the text encoding")


def text_to_numbers(table: pd.DataFrame) -> pd.DataFrame:
    """Turn columns that are clearly numbers into numbers. Anything doubtful stays text."""
    for column in table.columns:
        values = table[column].astype(str).str.strip()
        filled = values[values != ""]
        if len(filled) and filled.str.fullmatch(SAFE_NUMBER).all():
            table[column] = pd.to_numeric(values.where(values != ""))
    return table


def excel_types(table: pd.DataFrame) -> pd.DataFrame:
    """Columns where Excel stored only numbers (or only dates) become number (or date) columns."""
    for column in table.columns:
        filled = table[column].dropna()
        if not len(filled):
            continue
        if all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in filled):
            table[column] = pd.to_numeric(table[column])
        elif all(isinstance(v, (datetime, date)) for v in filled):
            table[column] = pd.to_datetime(table[column])
    return table


def read_json_file(path: Path) -> dict:
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except (json.JSONDecodeError, UnicodeDecodeError) as error:
        raise SheetProblem(f"it isn't valid JSON ({error})") from None

    def records_to_table(records):
        if all(isinstance(r, dict) for r in records):
            table = pd.json_normalize(records)  # nested objects become columns like "address.city"
        else:
            table = pd.DataFrame(records)
        # Lists inside a cell (like "tags": ["a", "b"]) are kept as JSON text.
        for column in table.columns:
            table[column] = table[column].map(lambda v: json.dumps(v, ensure_ascii=False)
                                              if isinstance(v, (list, dict)) else v)
        return table

    if isinstance(data, list):
        return {path.stem: records_to_table(data)}
    if isinstance(data, dict) and data and all(isinstance(v, list) for v in data.values()):
        if all(v and all(isinstance(r, dict) for r in v) for v in data.values()):
            # {"Sheet1": [...rows...], "Sheet2": [...]}, the way this script saves several sheets
            return {name: records_to_table(rows) for name, rows in data.items()}
        try:
            return {path.stem: pd.DataFrame(data)}  # {"column": [values], ...}
        except ValueError:
            raise SheetProblem("its lists have different lengths, so it can't be made into a table") from None
    if isinstance(data, dict):
        return {path.stem: records_to_table([data])}
    raise SheetProblem("it contains a single value, not a table")


def read_file(path: Path) -> dict:
    """Returns {sheet name: table}. CSV, TSV and JSON files have one sheet, named after the file."""
    ext = path.suffix.lower()
    try:
        if ext == ".csv":
            sheets = {path.stem: read_text_table(path)}
        elif ext in (".tsv", ".tab"):
            sheets = {path.stem: read_text_table(path, "\t")}
        elif ext == ".json":
            sheets = read_json_file(path)
        else:
            # calamine reads .xlsx, .xlsm, .xls and .ods, quickly. dtype=object keeps every cell
            # exactly as stored, so text like ZIP code "01234" isn't turned into the number 1234.
            sheets = pd.read_excel(path, sheet_name=None, engine="calamine", dtype=object)
            sheets = {name: excel_types(table) for name, table in sheets.items()}
    except SheetProblem:
        raise
    except ImportError as error:
        raise SheetProblem(f"a package is missing ({error}). Run:  pip install -r requirements.txt") from None
    except Exception:
        raise SheetProblem("it could not be read. The file may be damaged, or not really a spreadsheet") from None

    cleaned = {}
    for name, table in sheets.items():
        if table.empty and len(table.columns) == 0:
            continue  # completely empty sheet
        # Columns with no heading get "column_1", "column_2", ... instead of pandas' "Unnamed: 0".
        table.columns = [f"column_{int(m[1]) + 1}" if (m := re.fullmatch(r"Unnamed: (\d+)", str(c))) else str(c)
                         for c in table.columns]
        cleaned[str(name)] = whole_numbers(table)
    if not cleaned:
        raise SheetProblem("it's empty")
    return cleaned


def whole_numbers(table: pd.DataFrame) -> pd.DataFrame:
    """Columns of whole numbers with some blanks would show as 3.0, 4.0; keep them as 3, 4."""
    for column in table.columns:
        values = table[column]
        if pd.api.types.is_float_dtype(values):
            filled = values.dropna()
            if len(filled) and (filled % 1 == 0).all() and filled.abs().max() < 2 ** 53:
                table[column] = values.astype("Int64")
    return table


# ---------------------------------------------------------------- writing

def dates_as_text(table: pd.DataFrame) -> pd.DataFrame:
    """Dates as 2024-05-14 (or 2024-05-14 13:45:00 if they have a time), for text formats."""
    table = table.copy()
    for column in table.columns:
        values = table[column]
        if pd.api.types.is_datetime64_any_dtype(values):
            filled = values.dropna()
            has_time = bool((filled != filled.dt.normalize()).any())
            table[column] = values.dt.strftime("%Y-%m-%d %H:%M:%S" if has_time else "%Y-%m-%d")
    return table


def dates_for_excel(table: pd.DataFrame) -> pd.DataFrame:
    """Dates without a time are saved as plain dates, so Excel doesn't show 00:00:00."""
    table = table.copy()
    for column in table.columns:
        values = table[column]
        if pd.api.types.is_datetime64_any_dtype(values):
            filled = values.dropna()
            if len(filled) and (filled == filled.dt.normalize()).all():
                table[column] = values.dt.date
    return table


def safe_name(text: str) -> str:
    """Remove characters Windows doesn't allow in file names."""
    return re.sub(r'[<>:"/\\|?*]', "_", text).strip().rstrip(".") or "sheet"


def excel_sheet_names(names) -> list[str]:
    """Excel sheet names: at most 31 characters, no []:*?/\\ and no two the same."""
    result = []
    for name in names:
        clean = (re.sub(r"[\[\]:*?/\\]", "_", name).strip("'") or "Sheet")[:31]
        candidate, number = clean, 2
        while candidate.lower() in (r.lower() for r in result):
            suffix = f" ({number})"
            candidate = clean[:31 - len(suffix)] + suffix
            number += 1
        result.append(candidate)
    return result


def fit_column_widths(worksheet, table: pd.DataFrame):
    from openpyxl.utils import get_column_letter
    for i, column in enumerate(table.columns, start=1):
        longest = max([len(str(column))] + [len(str(v)) for v in table[column].head(1000) if pd.notna(v)])
        worksheet.column_dimensions[get_column_letter(i)].width = min(60, longest + 2)


def write_file(sheets: dict, fmt: str, out_base: Path) -> list[Path]:
    """Save all sheets in one format. Returns the files written."""
    written = []
    if fmt in ("csv", "tsv"):
        # CSV and TSV hold one sheet per file: book.xlsx with 2 sheets -> book_Sheet1.csv, book_Sheet2.csv
        for name, table in sheets.items():
            path = out_base.with_suffix(FORMATS[fmt]) if len(sheets) == 1 else \
                out_base.with_name(f"{out_base.name}_{safe_name(name)}{FORMATS[fmt]}")
            # utf-8-sig adds a marker that tells Excel the file is UTF-8, so accents show correctly.
            dates_as_text(table).to_csv(path, index=False, sep="," if fmt == "csv" else "\t", encoding="utf-8-sig")
            written.append(path)

    elif fmt in ("xlsx", "ods"):
        path = out_base.with_suffix(FORMATS[fmt])
        engine = "openpyxl" if fmt == "xlsx" else "odf"
        with pd.ExcelWriter(path, engine=engine) as writer:
            for sheet_name, table in zip(excel_sheet_names(sheets), sheets.values()):
                dates_for_excel(table).to_excel(writer, sheet_name=sheet_name, index=False)
                if fmt == "xlsx":
                    worksheet = writer.sheets[sheet_name]
                    fit_column_widths(worksheet, table)
                    worksheet.freeze_panes = "A2"  # keep the heading row visible while scrolling
        written.append(path)

    elif fmt == "json":
        path = out_base.with_suffix(".json")
        data = {name: json.loads(dates_as_text(table).to_json(orient="records", force_ascii=False))
                for name, table in sheets.items()}
        if len(data) == 1:
            data = next(iter(data.values()))  # one sheet: just the list of rows
        path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
        written.append(path)

    elif fmt == "html":
        path = out_base.with_suffix(".html")
        parts = []
        for name, table in sheets.items():
            if len(sheets) > 1:
                parts.append(f"<h2>{html.escape(name)}</h2>")
            parts.append(dates_as_text(table).to_html(index=False, na_rep="", border=0))
        path.write_text(HTML_PAGE.format(title=html.escape(out_base.name), body="\n".join(parts)), encoding="utf-8")
        written.append(path)
    return written


# ---------------------------------------------------------------- command line

def clean_formats(words: list[str]) -> list[str]:
    """Lowercase, drop leading dots and commas, and remove duplicates (keeping order)."""
    formats = []
    for word in words:
        for part in word.replace(",", " ").split():
            fmt = part.lower().lstrip(".")
            fmt = {"excel": "xlsx", "xls": "xlsx", "tab": "tsv", "htm": "html"}.get(fmt, fmt)
            if fmt not in formats:
                formats.append(fmt)
    return formats


def ask_formats() -> list[str]:
    """Keep asking until the user types valid formats."""
    print("\nAvailable formats: " + ", ".join(FORMATS))
    while True:
        answer = input("Which formats do you want? (e.g. 'xlsx csv', or 'all'): ").strip().lower()
        if answer == "all":
            return list(FORMATS)
        formats = clean_formats([answer])
        unknown = [f for f in formats if f not in FORMATS]
        if formats and not unknown:
            return formats
        print(f"  Unknown format(s): {', '.join(unknown)}. Try again." if unknown else "  Please type a format.")


def find_input_files(input_dir: Path) -> list[Path]:
    """Return the spreadsheets in input_dir. Other files are listed as ignored."""
    files = []
    for path in sorted(input_dir.iterdir(), key=lambda p: p.name.lower()):
        if not path.is_file() or path.name.startswith((".", "~$")):
            continue  # "~$" files are temporary files Excel makes while a workbook is open
        if path.suffix.lower() in INPUT_EXTENSIONS:
            files.append(path)
        else:
            print(f"  [ignored] {path.name} (not a supported spreadsheet)")
    return files


def describe(sheets: dict) -> str:
    rows = sum(len(t) for t in sheets.values())
    sheet_text = "1 sheet" if len(sheets) == 1 else f"{len(sheets)} sheets"
    return f"{sheet_text}, {rows:,} row{'s' if rows != 1 else ''}"


def main():
    parser = argparse.ArgumentParser(
        description="Convert every spreadsheet in the input folder into other formats.",
        epilog="Supported formats: " + ", ".join(FORMATS),
    )
    parser.add_argument("formats", nargs="*", help="formats to convert to, e.g. xlsx csv json")
    parser.add_argument("--all", action="store_true", help="convert to every supported format")
    parser.add_argument("-i", "--input", type=Path, default=INPUT_DIR,
                        help="folder to read spreadsheets from (default: the input folder)")
    parser.add_argument("-o", "--output", type=Path, default=OUTPUT_DIR,
                        help="folder to save converted files to (default: the output folder)")
    args = parser.parse_args()

    formats = list(FORMATS) if args.all else clean_formats(args.formats)
    unknown = [f for f in formats if f not in FORMATS]
    if unknown:
        sys.exit(f"Error: unknown format(s): {', '.join(unknown)}\nSupported: {', '.join(FORMATS)}")

    args.input.mkdir(parents=True, exist_ok=True)
    args.output.mkdir(parents=True, exist_ok=True)

    print(f"Looking in: {args.input}")
    files = find_input_files(args.input)
    if not files:
        print("No spreadsheets found. Put the files you want converted in that folder, then run this again.")
        return
    print(f"Found {len(files)} file(s):")
    for f in files:
        print(f"  - {f.name}")
    if not formats:
        formats = ask_formats()

    # "data.csv" and "data.xlsx" would both become "data.json",
    # so those keep their old extension in the name: "data_csv.json", "data_xlsx.json".
    stem_counts = Counter(f.stem.lower() for f in files)
    duplicate_stems = {stem for stem, count in stem_counts.items() if count > 1}

    converted = skipped = failed = 0
    for i, src in enumerate(files, start=1):
        print(f"\n[{i}/{len(files)}] {src.name}")
        try:
            sheets = read_file(src)
        except SheetProblem as problem:
            print(f"  [fail] {problem}")
            failed += 1
            continue
        print(f"  ({describe(sheets)})")

        stem = f"{src.stem}_{src.suffix.lstrip('.').lower()}" if src.stem.lower() in duplicate_stems else src.stem
        for fmt in formats:
            if INPUT_EXTENSIONS[src.suffix.lower()] == fmt:
                print(f"  [skip] {fmt:<5} (file is already {fmt})")
                skipped += 1
                continue
            fmt_dir = args.output / fmt
            fmt_dir.mkdir(exist_ok=True)
            try:
                written = write_file(sheets, fmt, fmt_dir / stem)
            except ImportError as error:
                print(f"  [fail] {fmt:<5} a package is missing ({error}). Run:  pip install -r requirements.txt")
                failed += 1
                continue
            except Exception as error:
                print(f"  [fail] {fmt:<5} {error}")
                failed += 1
                continue
            names = ", ".join(f"{fmt}/{p.name}" for p in written)
            print(f"  [ok]   {fmt:<5} -> {names}")
            converted += 1

    print(f"\nDone: {converted} converted, {skipped} skipped, {failed} failed.")
    print(f"Converted files are in: {args.output}")
    if failed:
        sys.exit(1)


if __name__ == "__main__":
    try:
        main()
    except (KeyboardInterrupt, EOFError):
        print("\nCancelled.")
        sys.exit(1)
