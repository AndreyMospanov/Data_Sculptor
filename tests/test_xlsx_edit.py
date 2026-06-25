from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook

from core.xlsx_csv import read_xlsx_rows
from core.xlsx_edit import (
    delete_column_in_xlsx,
    delete_column_in_xlsx_path,
    insert_column_in_xlsx,
    insert_column_in_xlsx_path,
    inspect_xlsx_columns,
    inspect_xlsx_path_columns,
    swap_columns_in_xlsx,
    swap_columns_in_xlsx_path,
)
def write_editable_xlsx(path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "First"
    worksheet.append(["Name", "Amount", None])
    worksheet.append(["Alice", None, 42])
    workbook.save(path)


class XlsxEditTests(unittest.TestCase):
    # Эти тесты работают с тем же минимальным XLSX, что и конвертер. Так мы
    # проверяем редактирование реального OOXML-архива без внешних библиотек.
    def test_inspects_header_columns_only(self) -> None:
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "input.xlsx"
            write_editable_xlsx(source)

            report = inspect_xlsx_columns(source)

            self.assertEqual(report.filled_column_count, 2)
            self.assertEqual([(column.letter, column.header, column.filled_cells) for column in report.columns], [
                ("A", "Name", 1),
                ("B", "Amount", 1),
            ])

    def test_inserts_column_with_header(self) -> None:
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "input.xlsx"
            write_editable_xlsx(source)

            result = insert_column_in_xlsx(source, "B", header="New")
            rows = read_xlsx_rows(source)

            self.assertEqual(result.column, "B")
            self.assertEqual(rows, [["Name", "New", "Amount"], ["Alice", "", "", "42"]])

    def test_inspects_folder(self) -> None:
        with TemporaryDirectory() as temp_dir:
            directory = Path(temp_dir)
            source_dir = directory / "xlsx"
            source_dir.mkdir()
            write_editable_xlsx(source_dir / "one.xlsx")
            write_editable_xlsx(source_dir / "two.xlsx")

            reports = inspect_xlsx_path_columns(source_dir)

            self.assertEqual([report.source.name for report in reports], ["one.xlsx", "two.xlsx"])
            self.assertEqual([report.filled_column_count for report in reports], [2, 2])

    def test_inserts_column_in_folder(self) -> None:
        with TemporaryDirectory() as temp_dir:
            directory = Path(temp_dir)
            source_dir = directory / "xlsx"
            source_dir.mkdir()
            one = source_dir / "one.xlsx"
            two = source_dir / "two.xlsx"
            write_editable_xlsx(one)
            write_editable_xlsx(two)

            results = insert_column_in_xlsx_path(source_dir, "A", header="First")

            self.assertEqual([result.source.name for result in results], ["one.xlsx", "two.xlsx"])
            self.assertEqual(read_xlsx_rows(one)[0], ["First", "Name", "Amount"])
            self.assertEqual(read_xlsx_rows(two)[0], ["First", "Name", "Amount"])

    def test_deletes_column(self) -> None:
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "input.xlsx"
            write_editable_xlsx(source)

            result = delete_column_in_xlsx(source, "B")
            rows = read_xlsx_rows(source)

            self.assertEqual(result.column, "B")
            self.assertEqual(rows, [["Name"], ["Alice", "42"]])

    def test_deletes_column_in_folder(self) -> None:
        with TemporaryDirectory() as temp_dir:
            directory = Path(temp_dir)
            source_dir = directory / "xlsx"
            source_dir.mkdir()
            one = source_dir / "one.xlsx"
            two = source_dir / "two.xlsx"
            write_editable_xlsx(one)
            write_editable_xlsx(two)

            results = delete_column_in_xlsx_path(source_dir, "A")

            self.assertEqual([result.source.name for result in results], ["one.xlsx", "two.xlsx"])
            self.assertEqual(read_xlsx_rows(one), [["Amount"], ["", "42"]])
            self.assertEqual(read_xlsx_rows(two), [["Amount"], ["", "42"]])

    def test_swaps_columns(self) -> None:
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "input.xlsx"
            write_editable_xlsx(source)

            result = swap_columns_in_xlsx(source, "A", "C")
            rows = read_xlsx_rows(source)

            self.assertEqual((result.first_column, result.second_column), ("A", "C"))
            self.assertEqual(rows, [["", "Amount", "Name"], ["42", "", "Alice"]])

    def test_swaps_columns_in_folder(self) -> None:
        with TemporaryDirectory() as temp_dir:
            directory = Path(temp_dir)
            source_dir = directory / "xlsx"
            source_dir.mkdir()
            one = source_dir / "one.xlsx"
            two = source_dir / "two.xlsx"
            write_editable_xlsx(one)
            write_editable_xlsx(two)

            results = swap_columns_in_xlsx_path(source_dir, "A", "B")

            self.assertEqual([result.source.name for result in results], ["one.xlsx", "two.xlsx"])
            self.assertEqual(read_xlsx_rows(one), [["Amount", "Name"], ["", "Alice", "42"]])
            self.assertEqual(read_xlsx_rows(two), [["Amount", "Name"], ["", "Alice", "42"]])


if __name__ == "__main__":
    unittest.main()
