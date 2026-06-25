from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from core.core import RenameError, apply_rename_plan, build_rename_plan


def touch(path: Path) -> None:
    # Тестам важны имена файлов, содержимое можно оставить пустым.
    path.write_text("", encoding="utf-8")


class BuildRenamePlanTests(unittest.TestCase):
    # Эти тесты покрывают ядро переименования без GUI и CLI: так проще ловить
    # ошибки в масках, датах и проверке конфликтов.
    def test_converts_date_and_swaps_name_with_date(self) -> None:
        with TemporaryDirectory() as temp_dir:
            directory = Path(temp_dir)
            touch(directory / "Report 31.05.2026.txt")

            plans = build_rename_plan(
                directory,
                "{name} {date:DD.MM.YYYY}.{ext}",
                "{date:YYYYMMDD}_{name}.{ext}",
            )

            self.assertEqual(
                [(p.source.name, p.target.name) for p in plans],
                [("Report 31.05.2026.txt", "20260531_Report.txt")],
            )

    def test_supports_output_name_before_date(self) -> None:
        with TemporaryDirectory() as temp_dir:
            directory = Path(temp_dir)
            touch(directory / "20260531_Report.txt")

            plans = build_rename_plan(
                directory,
                "{date:YYYYMMDD}_{name}.{ext}",
                "{name} {date:DD.MM.YYYY}.{ext}",
            )

            self.assertEqual(plans[0].target.name, "Report 31.05.2026.txt")

    def test_ignores_non_matching_files(self) -> None:
        with TemporaryDirectory() as temp_dir:
            directory = Path(temp_dir)
            touch(directory / "Report 31.05.2026.txt")
            touch(directory / "notes.txt")

            plans = build_rename_plan(
                directory,
                "{name} {date:DD.MM.YYYY}.{ext}",
                "{date:YYYYMMDD}_{name}.{ext}",
            )

            self.assertEqual(len(plans), 1)

    def test_rejects_invalid_date(self) -> None:
        with TemporaryDirectory() as temp_dir:
            directory = Path(temp_dir)
            touch(directory / "Report 99.05.2026.txt")

            with self.assertRaisesRegex(RenameError, "does not match format"):
                build_rename_plan(
                    directory,
                    "{name} {date:DD.MM.YYYY}.{ext}",
                    "{date:YYYYMMDD}_{name}.{ext}",
                )

    def test_rejects_target_conflicts(self) -> None:
        with TemporaryDirectory() as temp_dir:
            directory = Path(temp_dir)
            touch(directory / "Report 31.05.2026.txt")
            touch(directory / "Report 01.06.2026.txt")

            with self.assertRaisesRegex(RenameError, "Target name conflict"):
                build_rename_plan(
                    directory,
                    "{name} {date:DD.MM.YYYY}.{ext}",
                    "{name}.{ext}",
                )

    def test_applies_rename_plan(self) -> None:
        with TemporaryDirectory() as temp_dir:
            directory = Path(temp_dir)
            touch(directory / "Report 31.05.2026.txt")

            plans = build_rename_plan(
                directory,
                "{name} {date:DD.MM.YYYY}.{ext}",
                "{date:YYYYMMDD}_{name}.{ext}",
            )
            apply_rename_plan(plans)

            self.assertFalse((directory / "Report 31.05.2026.txt").exists())
            self.assertTrue((directory / "20260531_Report.txt").exists())


if __name__ == "__main__":
    unittest.main()
