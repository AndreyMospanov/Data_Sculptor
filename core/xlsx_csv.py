from __future__ import annotations

import csv
import re
import zipfile
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Callable, Iterable
from xml.etree import ElementTree


# Пространства имен OOXML. XLSX внутри является ZIP-архивом с XML-файлами,
# поэтому все обращения к workbook/sheetData/relationships идут через эти URI.
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"m": MAIN_NS, "r": REL_NS, "pr": PACKAGE_REL_NS}
CELL_REF_RE = re.compile(r"(?P<column>[A-Z]+)(?P<row>\d+)")
BUILTIN_DATE_FORMAT_IDS = {
    14,
    15,
    16,
    17,
    18,
    19,
    20,
    21,
    22,
    27,
    28,
    29,
    30,
    31,
    32,
    33,
    34,
    35,
    36,
    45,
    46,
    47,
    50,
    51,
    52,
    53,
    54,
    55,
    56,
    57,
    58,
}


class XlsxCsvError(Exception):
    """Raised when XLSX to CSV conversion cannot be completed."""


class ConversionCancelled(Exception):
    """Raised when a user-requested cancellation stops conversion."""


@dataclass(frozen=True)
class SheetInfo:
    # Имя листа показываем пользователю, а path нужен парсеру для чтения XML.
    name: str
    path: str


@dataclass(frozen=True)
class ConversionResult:
    source: Path
    target: Path


@dataclass(frozen=True)
class ConversionProgress:
    # Это маленький контракт между backend и GUI: backend ничего не знает о
    # tkinter, он только сообщает, какой файл и какая стадия сейчас выполняются.
    source: Path
    target: Path
    file_index: int
    file_count: int
    percent: int
    message: str


@dataclass(frozen=True)
class CellStyle:
    # Для CSV сейчас важен только один признак: является ли числовое значение
    # Excel-датой. Сам формат можно будет использовать позже для тонкой настройки.
    is_date: bool = False


ProgressCallback = Callable[[ConversionProgress], None]
CancelCallback = Callable[[], bool]


def convert_xlsx_to_csv(
    source: Path,
    target: Path,
    *,
    delimiter: str = ",",
    sheet: str | int | None = None,
    encoding: str = "utf-8-sig",
    progress: ProgressCallback | None = None,
    cancel_requested: CancelCallback | None = None,
    file_index: int = 1,
    file_count: int = 1,
) -> None:
    """Convert one XLSX worksheet to CSV."""

    if len(delimiter) != 1:
        raise XlsxCsvError("Delimiter must be exactly one character.")

    check_cancel(cancel_requested)
    emit_progress(progress, source, target, file_index, file_count, 0, "Starting")
    check_cancel(cancel_requested)
    emit_progress(progress, source, target, file_index, file_count, 35, "Reading worksheet")
    rows = read_xlsx_rows(source, sheet=sheet)
    check_cancel(cancel_requested)
    target.parent.mkdir(parents=True, exist_ok=True)
    emit_progress(progress, source, target, file_index, file_count, 85, "Writing CSV")
    check_cancel(cancel_requested)
    with target.open("w", newline="", encoding=encoding) as csv_file:
        writer = csv.writer(csv_file, delimiter=delimiter)
        writer.writerows(rows)
    emit_progress(progress, source, target, file_index, file_count, 100, "Done")


def convert_xlsx_path_to_csv(
    source: Path,
    target: Path,
    *,
    delimiter: str = ",",
    sheet: str | int | None = None,
    encoding: str = "utf-8-sig",
    recursive: bool = False,
    progress: ProgressCallback | None = None,
    cancel_requested: CancelCallback | None = None,
) -> list[ConversionResult]:
    """Convert a single XLSX file or all XLSX files in a folder."""

    source = Path(source)
    target = Path(target)
    if source.is_file():
        # Для одного файла caller сам задает CSV-путь. В GUI это режим
        # "выбран файл", в CLI это пара source.xlsx -> target.csv.
        convert_xlsx_to_csv(
            source,
            target,
            delimiter=delimiter,
            sheet=sheet,
            encoding=encoding,
            progress=progress,
            cancel_requested=cancel_requested,
            file_index=1,
            file_count=1,
        )
        return [ConversionResult(source=source, target=target)]

    if source.is_dir():
        if target.exists() and not target.is_dir():
            raise XlsxCsvError(f"Target must be a directory when source is a directory: {target}")

        # Список файлов собираем заранее: так GUI может показывать счетчик
        # "текущий файл / всего файлов", а backend сразу ловит пустую папку.
        xlsx_files = list(iter_xlsx_files(source, recursive=recursive))
        if not xlsx_files:
            raise XlsxCsvError(f"No .xlsx files found in: {source}")

        results: list[ConversionResult] = []
        for index, xlsx_file in enumerate(xlsx_files, start=1):
            check_cancel(cancel_requested)
            relative_parent = xlsx_file.parent.relative_to(source)
            csv_file = target / relative_parent / xlsx_file.with_suffix(".csv").name
            convert_xlsx_to_csv(
                xlsx_file,
                csv_file,
                delimiter=delimiter,
                sheet=sheet,
                encoding=encoding,
                progress=progress,
                cancel_requested=cancel_requested,
                file_index=index,
                file_count=len(xlsx_files),
            )
            results.append(ConversionResult(source=xlsx_file, target=csv_file))

        return results

    raise XlsxCsvError(f"Source must be an XLSX file or directory: {source}")


def iter_xlsx_files(directory: Path, *, recursive: bool) -> Iterable[Path]:
    # Временные файлы Excel вида "~$book.xlsx" пропускаем: они часто создаются
    # открытым Excel и не являются полноценными книгами для конвертации.
    pattern = "**/*.xlsx" if recursive else "*.xlsx"
    for path in sorted(directory.glob(pattern)):
        if path.is_file() and not path.name.startswith("~$"):
            yield path


def emit_progress(
    progress: ProgressCallback | None,
    source: Path,
    target: Path,
    file_index: int,
    file_count: int,
    percent: int,
    message: str,
) -> None:
    if progress is None:
        return
    progress(
        ConversionProgress(
            source=source,
            target=target,
            file_index=file_index,
            file_count=file_count,
            percent=percent,
            message=message,
        )
    )


def check_cancel(cancel_requested: CancelCallback | None) -> None:
    # Отмена мягкая: мы проверяем флаг между крупными стадиями, не прерывая
    # запись файла посередине системного вызова.
    if cancel_requested is not None and cancel_requested():
        raise ConversionCancelled("Conversion cancelled by user.")


def read_xlsx_rows(source: Path, *, sheet: str | int | None = None) -> list[list[str]]:
    source = Path(source)
    if not source.exists():
        raise XlsxCsvError(f"File does not exist: {source}")

    try:
        with zipfile.ZipFile(source) as archive:
            # Shared strings - общий словарь строк XLSX. Ячейка хранит индекс,
            # а настоящий текст лежит в xl/sharedStrings.xml.
            shared_strings = read_shared_strings(archive)
            styles = read_cell_styles(archive)
            sheet_info = choose_sheet(read_sheets(archive), sheet)
            return read_sheet_rows(archive, sheet_info.path, shared_strings, styles)
    except zipfile.BadZipFile as exc:
        raise XlsxCsvError(f"Not a valid XLSX file: {source}") from exc
    except KeyError as exc:
        raise XlsxCsvError(f"XLSX is missing required part: {exc}") from exc
    except ElementTree.ParseError as exc:
        raise XlsxCsvError(f"XLSX contains invalid XML: {exc}") from exc


def list_xlsx_sheets(source: Path) -> list[str]:
    try:
        with zipfile.ZipFile(source) as archive:
            return [sheet.name for sheet in read_sheets(archive)]
    except zipfile.BadZipFile as exc:
        raise XlsxCsvError(f"Not a valid XLSX file: {source}") from exc


def read_sheets(archive: zipfile.ZipFile) -> list[SheetInfo]:
    # workbook.xml содержит имена листов, а workbook.xml.rels связывает r:id
    # с физическим XML-файлом листа внутри архива.
    workbook = ElementTree.fromstring(archive.read("xl/workbook.xml"))
    relationships = read_relationships(archive, "xl/_rels/workbook.xml.rels")
    sheets: list[SheetInfo] = []

    for sheet in workbook.findall("m:sheets/m:sheet", NS):
        name = sheet.attrib["name"]
        relationship_id = sheet.attrib[f"{{{REL_NS}}}id"]
        target = relationships[relationship_id]
        sheets.append(SheetInfo(name=name, path=normalize_xl_path(target)))

    if not sheets:
        raise XlsxCsvError("Workbook does not contain sheets.")
    return sheets


def read_relationships(archive: zipfile.ZipFile, path: str) -> dict[str, str]:
    rels = ElementTree.fromstring(archive.read(path))
    result: dict[str, str] = {}
    for rel in rels.findall("pr:Relationship", NS):
        result[rel.attrib["Id"]] = rel.attrib["Target"]
    return result


def read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    # Файл sharedStrings.xml необязателен: небольшие XLSX могут хранить строки
    # прямо внутри ячеек как inlineStr.
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []

    root = ElementTree.fromstring(archive.read("xl/sharedStrings.xml"))
    return [read_text(si) for si in root.findall("m:si", NS)]


def read_cell_styles(archive: zipfile.ZipFile) -> list[CellStyle]:
    # styles.xml хранит форматирование ячеек. Для дат Excel хранит в ячейке
    # число, а стиль отдельно говорит: "показывать это число как дату".
    if "xl/styles.xml" not in archive.namelist():
        return []

    root = ElementTree.fromstring(archive.read("xl/styles.xml"))
    custom_formats: dict[int, str] = {}
    for num_format in root.findall("m:numFmts/m:numFmt", NS):
        custom_formats[int(num_format.attrib["numFmtId"])] = num_format.attrib.get("formatCode", "")

    styles: list[CellStyle] = []
    for xf in root.findall("m:cellXfs/m:xf", NS):
        num_format_id = int(xf.attrib.get("numFmtId", "0"))
        styles.append(CellStyle(is_date=is_date_num_format(num_format_id, custom_formats.get(num_format_id, ""))))
    return styles


def is_date_num_format(num_format_id: int, format_code: str) -> bool:
    if num_format_id in BUILTIN_DATE_FORMAT_IDS:
        return True
    if not format_code:
        return False

    cleaned = strip_format_literals(format_code).lower()
    has_date_token = any(token in cleaned for token in ("yy", "yyyy", "dd", "mmm"))
    has_time_token = "h" in cleaned and ("s" in cleaned or "m" in cleaned)
    return has_date_token or has_time_token


def strip_format_literals(format_code: str) -> str:
    # Убираем литералы и цветовые/условные секции Excel-формата, чтобы текст
    # внутри кавычек или [] не был ошибочно принят за date-токен.
    without_quotes = re.sub(r'"[^"]*"', "", format_code)
    without_brackets = re.sub(r"\[[^\]]*\]", "", without_quotes)
    return re.sub(r"\\.", "", without_brackets)


def read_sheet_rows(
    archive: zipfile.ZipFile,
    path: str,
    shared_strings: list[str],
    styles: list[CellStyle],
) -> list[list[str]]:
    root = ElementTree.fromstring(archive.read(path))
    rows: list[list[str]] = []

    for row in root.findall("m:sheetData/m:row", NS):
        # XLSX не обязан явно хранить пустые ячейки. Поэтому собираем значения
        # по индексу колонки и потом восстанавливаем пропуски пустыми строками.
        values_by_index: dict[int, str] = {}
        for cell in row.findall("m:c", NS):
            column_index = cell_column_index(cell)
            values_by_index[column_index] = read_cell_value(cell, shared_strings, styles)

        if values_by_index:
            max_index = max(values_by_index)
            rows.append([values_by_index.get(index, "") for index in range(max_index + 1)])
        else:
            rows.append([])

    return trim_trailing_empty_rows(rows)


def read_cell_value(cell: ElementTree.Element, shared_strings: list[str], styles: list[CellStyle]) -> str:
    # Поддерживаем основные типы, которые нужны для CSV: shared string,
    # inline string, boolean и обычное числовое/текстовое значение из <v>.
    cell_type = cell.attrib.get("t")
    if cell_type == "s":
        index_text = read_value_node(cell)
        if index_text == "":
            return ""
        return shared_strings[int(index_text)]
    if cell_type == "inlineStr":
        inline = cell.find("m:is", NS)
        return read_text(inline) if inline is not None else ""
    if cell_type == "b":
        return "TRUE" if read_value_node(cell) == "1" else "FALSE"

    value = read_value_node(cell)
    if value and cell_has_date_style(cell, styles):
        return format_excel_date(value)
    return value


def cell_has_date_style(cell: ElementTree.Element, styles: list[CellStyle]) -> bool:
    style_text = cell.attrib.get("s")
    if style_text is None:
        return False
    try:
        style_index = int(style_text)
    except ValueError:
        return False
    return 0 <= style_index < len(styles) and styles[style_index].is_date


def format_excel_date(value: str) -> str:
    try:
        serial = float(value)
    except ValueError:
        return value

    # Стандартная Excel date system 1900: 44953 -> 27.01.2023.
    moment = datetime(1899, 12, 30) + timedelta(days=serial)
    if moment.time() == datetime.min.time():
        return moment.strftime("%d.%m.%Y")
    return moment.strftime("%d.%m.%Y %H:%M:%S")


def read_value_node(cell: ElementTree.Element) -> str:
    value = cell.find("m:v", NS)
    return value.text if value is not None and value.text is not None else ""


def read_text(element: ElementTree.Element | None) -> str:
    if element is None:
        return ""
    return "".join(text_node.text or "" for text_node in element.findall(".//m:t", NS))


def choose_sheet(sheets: list[SheetInfo], sheet: str | int | None) -> SheetInfo:
    # Пользователь может выбрать лист по имени или по 1-based индексу, как это
    # обычно видно в интерфейсе Excel.
    if sheet is None:
        return sheets[0]

    if isinstance(sheet, int) or str(sheet).isdigit():
        index = int(sheet) - 1
        if 0 <= index < len(sheets):
            return sheets[index]
        raise XlsxCsvError(f"Sheet index is out of range: {sheet}")

    for sheet_info in sheets:
        if sheet_info.name == sheet:
            return sheet_info
    raise XlsxCsvError(f"Sheet not found: {sheet}")


def cell_column_index(cell: ElementTree.Element) -> int:
    reference = cell.attrib.get("r")
    if not reference:
        return 0

    match = CELL_REF_RE.fullmatch(reference)
    if not match:
        raise XlsxCsvError(f"Invalid cell reference: {reference}")
    return column_name_to_index(match.group("column"))


def column_name_to_index(column: str) -> int:
    # Excel-колонки A, B, ..., Z, AA переводим в zero-based индекс для списка.
    index = 0
    for char in column:
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index - 1


def normalize_xl_path(target: str) -> str:
    # Relationship target может быть "worksheets/sheet1.xml" или уже
    # "xl/worksheets/sheet1.xml"; приводим оба варианта к пути внутри архива.
    normalized = target.lstrip("/")
    if normalized.startswith("xl/"):
        return normalized
    return f"xl/{normalized}"


def trim_trailing_empty_rows(rows: Iterable[list[str]]) -> list[list[str]]:
    # Последние полностью пустые строки не несут данных и в CSV только создают
    # визуальный шум, поэтому отрезаем их.
    result = list(rows)
    while result and all(value == "" for value in result[-1]):
        result.pop()
    return result
