from __future__ import annotations

import csv
import shutil
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable


class CsvEditError(Exception):
    """Ошибка анализа или редактирования CSV."""


@dataclass(frozen=True)
class CsvColumnInfo:
    letter: str
    index: int
    header: str


@dataclass(frozen=True)
class CsvColumnReport:
    source: Path
    column_count: int
    columns: list[CsvColumnInfo]


@dataclass(frozen=True)
class CsvEditResult:
    source: Path
    operation: str


def inspect_csv_columns(
    source: Path,
    *,
    delimiter: str = ";",
    encoding: str = "utf-8-sig",
) -> CsvColumnReport:
    """Быстро прочитать только первую строку CSV с заголовками."""

    validate_delimiter(delimiter)
    source = Path(source)
    try:
        with source.open("r", newline="", encoding=encoding) as csv_file:
            header = next(csv.reader(csv_file, delimiter=delimiter), [])
    except OSError as exc:
        raise CsvEditError(f"Cannot read CSV: {exc}") from exc

    columns = [
        CsvColumnInfo(letter=index_to_column_name(index), index=index, header=value)
        for index, value in enumerate(header)
        if value != ""
    ]
    return CsvColumnReport(source=source, column_count=len(columns), columns=columns)


def inspect_csv_path_columns(
    source: Path,
    *,
    delimiter: str = ";",
    encoding: str = "utf-8-sig",
    recursive: bool = False,
) -> list[CsvColumnReport]:
    return [
        inspect_csv_columns(path, delimiter=delimiter, encoding=encoding)
        for path in resolve_csv_files(source, recursive=recursive)
    ]


def insert_column_in_csv(
    source: Path,
    column: str,
    *,
    header: str = "",
    delimiter: str = ";",
    encoding: str = "utf-8-sig",
) -> CsvEditResult:
    index = column_name_to_index(column)
    rewrite_csv(
        Path(source),
        delimiter=delimiter,
        encoding=encoding,
        transform=lambda row, row_index: insert_value(row, index, header if row_index == 0 else ""),
    )
    return CsvEditResult(source=Path(source), operation="insert")


def delete_column_in_csv(
    source: Path,
    column: str,
    *,
    delimiter: str = ";",
    encoding: str = "utf-8-sig",
) -> CsvEditResult:
    index = column_name_to_index(column)
    rewrite_csv(
        Path(source),
        delimiter=delimiter,
        encoding=encoding,
        transform=lambda row, _row_index: delete_value(row, index),
    )
    return CsvEditResult(source=Path(source), operation="delete")


def swap_columns_in_csv(
    source: Path,
    first_column: str,
    second_column: str,
    *,
    delimiter: str = ";",
    encoding: str = "utf-8-sig",
) -> CsvEditResult:
    first_index = column_name_to_index(first_column)
    second_index = column_name_to_index(second_column)
    if first_index == second_index:
        raise CsvEditError("Columns to swap must be different.")
    rewrite_csv(
        Path(source),
        delimiter=delimiter,
        encoding=encoding,
        transform=lambda row, _row_index: swap_values(row, first_index, second_index),
    )
    return CsvEditResult(source=Path(source), operation="swap")


def rewrite_csv(source: Path, *, delimiter: str, encoding: str, transform) -> None:
    """Переписать CSV во временный файл и заменить исходник только после успеха."""

    validate_delimiter(delimiter)
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".csv", dir=source.parent) as temp_file:
            temp_path = Path(temp_file.name)

        try:
            with source.open("r", newline="", encoding=encoding) as source_file, temp_path.open(
                "w", newline="", encoding=encoding
            ) as target_file:
                reader = csv.reader(source_file, delimiter=delimiter)
                writer = csv.writer(target_file, delimiter=delimiter)
                for row_index, row in enumerate(reader):
                    writer.writerow(transform(row, row_index))
            shutil.move(str(temp_path), source)
        finally:
            if temp_path.exists():
                temp_path.unlink()
    except OSError as exc:
        raise CsvEditError(f"Cannot edit CSV: {exc}") from exc


def insert_value(row: list[str], index: int, value: str) -> list[str]:
    result = list(row)
    while len(result) < index:
        result.append("")
    result.insert(index, value)
    return result


def delete_value(row: list[str], index: int) -> list[str]:
    result = list(row)
    if index < len(result):
        result.pop(index)
    return result


def swap_values(row: list[str], first_index: int, second_index: int) -> list[str]:
    result = list(row)
    while len(result) <= max(first_index, second_index):
        result.append("")
    result[first_index], result[second_index] = result[second_index], result[first_index]
    return result


def resolve_csv_files(source: Path, *, recursive: bool) -> list[Path]:
    source = Path(source)
    if source.is_file():
        return [source]
    if source.is_dir():
        files = list(iter_csv_files(source, recursive=recursive))
        if files:
            return files
        raise CsvEditError(f"No .csv files found in: {source}")
    raise CsvEditError(f"Source must be a CSV file or directory: {source}")


def iter_csv_files(directory: Path, *, recursive: bool) -> Iterable[Path]:
    pattern = "**/*.csv" if recursive else "*.csv"
    yield from (path for path in sorted(directory.glob(pattern)) if path.is_file())


def validate_delimiter(delimiter: str) -> None:
    if len(delimiter) != 1:
        raise CsvEditError("Delimiter must be exactly one character.")


def column_name_to_index(column: str) -> int:
    cleaned = column.strip().upper()
    if not cleaned or not cleaned.isalpha():
        raise CsvEditError(f"Invalid column name: {column}")
    index = 0
    for char in cleaned:
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index - 1


def index_to_column_name(index: int) -> str:
    if index < 0:
        raise CsvEditError(f"Invalid column index: {index}")
    result = ""
    value = index + 1
    while value:
        value, remainder = divmod(value - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result
