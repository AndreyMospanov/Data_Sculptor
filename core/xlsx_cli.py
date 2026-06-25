from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .xlsx_csv import XlsxCsvError, convert_xlsx_path_to_csv, list_xlsx_sheets


def build_parser() -> argparse.ArgumentParser:
    # CLI-конвертер умеет работать и с одним XLSX, и с папкой. GUI использует
    # тот же backend, поэтому поведение двух интерфейсов остается одинаковым.
    parser = argparse.ArgumentParser(
        prog="xlsx-to-csv",
        description="Convert XLSX worksheet data to CSV.",
    )
    parser.add_argument("source", type=Path, help="Input .xlsx file or directory with .xlsx files.")
    parser.add_argument("target", type=Path, nargs="?", help="Output .csv file or output directory.")
    parser.add_argument(
        "--delimiter",
        "-d",
        default=",",
        help="CSV delimiter, for example ',' ';' or tab. Default: ','.",
    )
    parser.add_argument("--sheet", "-s", help="Sheet name or 1-based sheet index. Default: first sheet.")
    parser.add_argument("--encoding", default="utf-8-sig", help="Output encoding. Default: utf-8-sig.")
    parser.add_argument("--list-sheets", action="store_true", help="Print workbook sheet names and exit.")
    parser.add_argument("--recursive", "-r", action="store_true", help="Process nested folders when source is a directory.")
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    # В командной строке табуляцию удобнее передавать как строку "\t", а csv
    # writer ожидает реальный символ табуляции.
    delimiter = "\t" if args.delimiter == "\\t" else args.delimiter

    try:
        if args.list_sheets:
            # Просмотр листов нужен перед выбором --sheet, особенно когда книга
            # пришла от внешней системы с неожиданными именами вкладок.
            for index, name in enumerate(list_xlsx_sheets(args.source), start=1):
                print(f"{index}: {name}")
            return 0

        if args.target is None:
            parser.error("target is required unless --list-sheets is used")

        # convert_xlsx_path_to_csv сам определяет, source - файл или папка.
        results = convert_xlsx_path_to_csv(
            args.source,
            args.target,
            delimiter=delimiter,
            sheet=args.sheet,
            encoding=args.encoding,
            recursive=args.recursive,
        )
    except XlsxCsvError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2

    for result in results:
        # Печатаем все пары source->target: это простой лог пакетной операции.
        print(f"{result.source} -> {result.target}")
    print(f"Converted: {len(results)} file(s)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
