"""Публичный backend API приложения.

Здесь собраны функции и классы, которые можно импортировать из GUI, CLI или
будущих интеграций. Внутренние helper-функции остаются в своих модулях.
"""

from .core import RenameError, RenamePlan, build_rename_plan
from .xlsx_csv import (
    ConversionCancelled,
    ConversionProgress,
    ConversionResult,
    XlsxCsvError,
    convert_xlsx_path_to_csv,
    convert_xlsx_to_csv,
    list_xlsx_sheets,
    read_xlsx_rows,
)
from .csv_edit import (
    CsvColumnInfo,
    CsvColumnReport,
    CsvEditError,
    CsvEditResult,
    delete_column_in_csv,
    insert_column_in_csv,
    inspect_csv_path_columns,
    swap_columns_in_csv,
)

_XLSX_EDIT_EXPORTS = {
    "ColumnInfo",
    "ColumnDeleteResult",
    "ColumnInsertResult",
    "ColumnSwapResult",
    "WorkbookColumnReport",
    "XlsxEditError",
    "delete_column_in_xlsx_path",
    "insert_column_in_xlsx_path",
    "inspect_xlsx_path_columns",
    "swap_columns_in_xlsx_path",
}

__all__ = [
    # Явный __all__ фиксирует публичную поверхность пакета и помогает не
    # привязывать внешние интерфейсы к внутренним деталям реализации.
    "RenameError",
    "RenamePlan",
    "ConversionProgress",
    "ConversionCancelled",
    "ConversionResult",
    "XlsxCsvError",
    "build_rename_plan",
    "convert_xlsx_path_to_csv",
    "convert_xlsx_to_csv",
    "list_xlsx_sheets",
    "read_xlsx_rows",
    "CsvColumnInfo",
    "CsvColumnReport",
    "CsvEditError",
    "CsvEditResult",
    "delete_column_in_csv",
    "insert_column_in_csv",
    "inspect_csv_path_columns",
    "swap_columns_in_csv",
    "ColumnInfo",
    "ColumnDeleteResult",
    "ColumnInsertResult",
    "ColumnSwapResult",
    "WorkbookColumnReport",
    "XlsxEditError",
    "delete_column_in_xlsx_path",
    "insert_column_in_xlsx_path",
    "inspect_xlsx_path_columns",
    "swap_columns_in_xlsx_path",
]


def __getattr__(name: str):
    # XLSX editor depends on openpyxl. Import it only when the caller asks for
    # editor-specific API, so the GUI and other features can start without it.
    if name in _XLSX_EDIT_EXPORTS:
        from . import xlsx_edit

        return getattr(xlsx_edit, name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
