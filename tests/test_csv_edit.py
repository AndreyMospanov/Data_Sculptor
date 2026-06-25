from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from core.csv_edit import (
    delete_column_in_csv,
    insert_column_in_csv,
    inspect_csv_columns,
    swap_columns_in_csv,
)


def write_csv(path: Path) -> None:
    path.write_text("Name;Amount;Date\nAlice;10;2023-01-01\nBob;20;2023-01-02\n", encoding="utf-8")


class CsvEditTests(unittest.TestCase):
    def test_inspects_headers(self) -> None:
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "input.csv"
            write_csv(source)

            report = inspect_csv_columns(source)

            self.assertEqual(report.column_count, 3)
            self.assertEqual([(column.letter, column.header) for column in report.columns], [("A", "Name"), ("B", "Amount"), ("C", "Date")])

    def test_inserts_column(self) -> None:
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "input.csv"
            write_csv(source)

            insert_column_in_csv(source, "B", header="Code")

            self.assertEqual(source.read_text(encoding="utf-8-sig"), "Name;Code;Amount;Date\nAlice;;10;2023-01-01\nBob;;20;2023-01-02\n")

    def test_deletes_column(self) -> None:
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "input.csv"
            write_csv(source)

            delete_column_in_csv(source, "B")

            self.assertEqual(source.read_text(encoding="utf-8-sig"), "Name;Date\nAlice;2023-01-01\nBob;2023-01-02\n")

    def test_swaps_columns(self) -> None:
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "input.csv"
            write_csv(source)

            swap_columns_in_csv(source, "A", "C")

            self.assertEqual(source.read_text(encoding="utf-8-sig"), "Date;Amount;Name\n2023-01-01;10;Alice\n2023-01-02;20;Bob\n")


if __name__ == "__main__":
    unittest.main()
