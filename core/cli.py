from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .core import RenameError, apply_rename_plan, build_rename_plan


def build_parser() -> argparse.ArgumentParser:
    # CLI переименования оставлен отдельным от GUI: это удобно для автоматизации
    # и одновременно дает стабильный backend-контракт для будущих интерфейсов.
    parser = argparse.ArgumentParser(
        prog="file-renamer",
        description="Batch rename files with input and output masks.",
    )
    parser.add_argument("directory", type=Path, help="Directory with files to rename.")
    parser.add_argument("--input", "-i", required=True, help="Input mask, e.g. '{name}_{date:DD.MM.YYYY}.{ext}'.")
    parser.add_argument("--output", "-o", required=True, help="Output mask, e.g. '{date:YYYYMMDD}_{name}.{ext}'.")
    parser.add_argument("--apply", action="store_true", help="Actually rename files. Default is preview only.")
    parser.add_argument("--recursive", "-r", action="store_true", help="Process nested directories.")
    parser.add_argument("--include-dirs", action="store_true", help="Rename directories too.")
    parser.add_argument("--overwrite", action="store_true", help="Allow replacing existing target files.")
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        # Сначала всегда строим план. Даже при --apply это дает возможность
        # заранее поймать конфликты имен и не выполнить половину операции.
        plans = build_rename_plan(
            args.directory,
            args.input,
            args.output,
            recursive=args.recursive,
            include_dirs=args.include_dirs,
            overwrite=args.overwrite,
        )
    except RenameError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2

    if not plans:
        print("No matching files.")
        return 0

    for plan in plans:
        # Preview и apply печатают один и тот же план, чтобы пользователь видел,
        # что именно будет или уже было переименовано.
        print(f"{plan.source.name} -> {plan.target.name}")

    if args.apply:
        apply_rename_plan(plans)
        print(f"Renamed: {len(plans)}")
    else:
        print(f"Preview only: {len(plans)} item(s). Add --apply to rename.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
