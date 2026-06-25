from pathlib import Path
from tempfile import TemporaryDirectory
import zipfile
import unittest

from core.xlsx_csv import (
    ConversionCancelled,
    convert_xlsx_path_to_csv,
    convert_xlsx_to_csv,
    list_xlsx_sheets,
    read_xlsx_rows,
)


def write_minimal_xlsx(path: Path) -> None:
    # Создаем минимальную XLSX-книгу вручную как ZIP с XML-файлами. Это держит
    # тесты независимыми от openpyxl/pandas и повторяет реальный формат XLSX.
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr(
            "xl/workbook.xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="First" sheetId="1" r:id="rId1"/>
    <sheet name="Second" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>""",
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Target="worksheets/sheet2.xml"/>
</Relationships>""",
        )
        archive.writestr(
            "xl/sharedStrings.xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <si><t>Name</t></si>
  <si><t>Amount</t></si>
  <si><t>Alice</t></si>
  <si><t>Bob</t></si>
</sst>""",
        )
        archive.writestr(
            "xl/worksheets/sheet1.xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>2</v></c>
      <c r="C2"><v>42</v></c>
    </row>
  </sheetData>
</worksheet>""",
        )
        archive.writestr(
            "xl/worksheets/sheet2.xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Code</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Enabled</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>X-1</t></is></c>
      <c r="B2" t="b"><v>1</v></c>
    </row>
  </sheetData>
</worksheet>""",
        )


def write_date_xlsx(path: Path) -> None:
    # В этом XLSX дата хранится как число 44953, а стиль s="1" указывает,
    # что это дата. Именно так Excel часто хранит даты внутри файла.
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr(
            "xl/workbook.xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Dates" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>""",
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Target="worksheets/sheet1.xml"/>
</Relationships>""",
        )
        archive.writestr(
            "xl/styles.xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cellXfs count="2">
    <xf numFmtId="0"/>
    <xf numFmtId="14"/>
  </cellXfs>
</styleSheet>""",
        )
        archive.writestr(
            "xl/worksheets/sheet1.xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Date</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Raw number</t></is></c>
    </row>
    <row r="2">
      <c r="A2" s="1"><v>44953</v></c>
      <c r="B2"><v>44953</v></c>
    </row>
  </sheetData>
</worksheet>""",
        )


class XlsxCsvTests(unittest.TestCase):
    # Тесты проверяют не Excel как приложение, а наш маленький OOXML-парсер:
    # shared strings, inline strings, boolean, папки и события прогресса.
    def test_reads_first_sheet_by_default(self) -> None:
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "input.xlsx"
            write_minimal_xlsx(source)

            rows = read_xlsx_rows(source)

            self.assertEqual(rows, [["Name", "Amount"], ["Alice", "", "42"]])

    def test_reads_sheet_by_name(self) -> None:
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "input.xlsx"
            write_minimal_xlsx(source)

            rows = read_xlsx_rows(source, sheet="Second")

            self.assertEqual(rows, [["Code", "Enabled"], ["X-1", "TRUE"]])

    def test_lists_sheet_names(self) -> None:
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "input.xlsx"
            write_minimal_xlsx(source)

            self.assertEqual(list_xlsx_sheets(source), ["First", "Second"])

    def test_converts_to_csv_with_custom_delimiter(self) -> None:
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "input.xlsx"
            target = Path(temp_dir) / "output.csv"
            write_minimal_xlsx(source)

            convert_xlsx_to_csv(source, target, delimiter=";", sheet=2, encoding="utf-8")

            self.assertEqual(target.read_text(encoding="utf-8"), "Code;Enabled\nX-1;TRUE\n")

    def test_converts_excel_serial_date_when_cell_has_date_style(self) -> None:
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "dates.xlsx"
            target = Path(temp_dir) / "dates.csv"
            write_date_xlsx(source)

            convert_xlsx_to_csv(source, target, delimiter=";", encoding="utf-8")

            self.assertEqual(target.read_text(encoding="utf-8"), "Date;Raw number\n27.01.2023;44953\n")

    def test_converts_folder_to_csv_files(self) -> None:
        with TemporaryDirectory() as temp_dir:
            directory = Path(temp_dir)
            source_dir = directory / "xlsx"
            target_dir = directory / "csv"
            source_dir.mkdir()
            write_minimal_xlsx(source_dir / "one.xlsx")
            write_minimal_xlsx(source_dir / "two.xlsx")

            results = convert_xlsx_path_to_csv(source_dir, target_dir, delimiter=";", sheet="Second", encoding="utf-8")

            self.assertEqual([result.target.name for result in results], ["one.csv", "two.csv"])
            self.assertEqual((target_dir / "one.csv").read_text(encoding="utf-8"), "Code;Enabled\nX-1;TRUE\n")
            self.assertEqual((target_dir / "two.csv").read_text(encoding="utf-8"), "Code;Enabled\nX-1;TRUE\n")

    def test_converts_folder_recursively(self) -> None:
        with TemporaryDirectory() as temp_dir:
            directory = Path(temp_dir)
            source_dir = directory / "xlsx"
            nested_dir = source_dir / "nested"
            target_dir = directory / "csv"
            nested_dir.mkdir(parents=True)
            write_minimal_xlsx(nested_dir / "inside.xlsx")

            results = convert_xlsx_path_to_csv(source_dir, target_dir, delimiter=",", encoding="utf-8", recursive=True)

            self.assertEqual(len(results), 1)
            self.assertTrue((target_dir / "nested" / "inside.csv").exists())

    def test_reports_progress_for_each_file(self) -> None:
        with TemporaryDirectory() as temp_dir:
            directory = Path(temp_dir)
            source_dir = directory / "xlsx"
            target_dir = directory / "csv"
            source_dir.mkdir()
            write_minimal_xlsx(source_dir / "one.xlsx")
            write_minimal_xlsx(source_dir / "two.xlsx")
            events = []

            convert_xlsx_path_to_csv(source_dir, target_dir, encoding="utf-8", progress=events.append)

            self.assertEqual([event.percent for event in events if event.message == "Starting"], [0, 0])
            self.assertEqual([event.percent for event in events if event.message == "Done"], [100, 100])
            self.assertEqual(events[-1].file_index, 2)
            self.assertEqual(events[-1].file_count, 2)

    def test_can_cancel_conversion(self) -> None:
        with TemporaryDirectory() as temp_dir:
            directory = Path(temp_dir)
            source = directory / "input.xlsx"
            target = directory / "output.csv"
            write_minimal_xlsx(source)

            with self.assertRaises(ConversionCancelled):
                convert_xlsx_path_to_csv(source, target, cancel_requested=lambda: True)

            self.assertFalse(target.exists())


if __name__ == "__main__":
    unittest.main()
