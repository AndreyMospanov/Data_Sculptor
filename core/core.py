from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable


PLACEHOLDER_RE = re.compile(r"\{(?P<name>[A-Za-z_][A-Za-z0-9_]*)(?::(?P<format>[^{}]+))?\}")
DATE_TOKENS = ("YYYY", "YY", "MM", "DD")


class RenameError(Exception):
    """Raised when masks or rename plans are invalid."""


@dataclass(frozen=True)
class FieldSpec:
    name: str
    format: str | None = None


@dataclass(frozen=True)
class RenamePlan:
    source: Path
    target: Path


@dataclass(frozen=True)
class CompiledInputMask:
    regex: re.Pattern[str]
    fields: dict[str, FieldSpec]


def build_rename_plan(
    directory: Path,
    input_mask: str,
    output_mask: str,
    *,
    recursive: bool = False,
    include_dirs: bool = False,
    overwrite: bool = False,
) -> list[RenamePlan]:
    """Return validated source-target pairs for files matching the input mask."""

    directory = Path(directory)
    if not directory.exists():
        raise RenameError(f"Directory does not exist: {directory}")
    if not directory.is_dir():
        raise RenameError(f"Path is not a directory: {directory}")

    compiled = compile_input_mask(input_mask)
    plans: list[RenamePlan] = []

    for path in iter_paths(directory, recursive=recursive, include_dirs=include_dirs):
        match = compiled.regex.fullmatch(path.name)
        if not match:
            continue

        values = normalize_values(match.groupdict(), compiled.fields)
        target_name = render_output_mask(output_mask, values, compiled.fields)
        target = path.with_name(target_name)
        if target == path:
            continue
        plans.append(RenamePlan(source=path, target=target))

    validate_plan(plans, overwrite=overwrite)
    return plans


def apply_rename_plan(plans: Iterable[RenamePlan]) -> None:
    for plan in plans:
        plan.source.rename(plan.target)


def compile_input_mask(mask: str) -> CompiledInputMask:
    fields: dict[str, FieldSpec] = {}
    regex_parts: list[str] = []
    position = 0

    for match in PLACEHOLDER_RE.finditer(mask):
        regex_parts.append(re.escape(mask[position : match.start()]))
        spec = FieldSpec(name=match.group("name"), format=match.group("format"))
        if spec.name in fields:
            raise RenameError(f"Duplicate field in input mask: {spec.name}")
        fields[spec.name] = spec
        regex_parts.append(f"(?P<{spec.name}>{regex_for_spec(spec)})")
        position = match.end()

    regex_parts.append(re.escape(mask[position:]))
    return CompiledInputMask(regex=re.compile("".join(regex_parts)), fields=fields)


def iter_paths(directory: Path, *, recursive: bool, include_dirs: bool) -> Iterable[Path]:
    paths = directory.rglob("*") if recursive else directory.iterdir()
    for path in paths:
        if path.is_file() or (include_dirs and path.is_dir()):
            yield path


def normalize_values(raw_values: dict[str, str], fields: dict[str, FieldSpec]) -> dict[str, object]:
    values: dict[str, object] = {}
    for name, value in raw_values.items():
        spec = fields[name]
        if spec.format and is_date_format(spec.format):
            values[name] = parse_date(value, spec.format)
        else:
            values[name] = value
    return values


def render_output_mask(mask: str, values: dict[str, object], fields: dict[str, FieldSpec]) -> str:
    def replace(match: re.Match[str]) -> str:
        name = match.group("name")
        fmt = match.group("format")
        if name not in values:
            raise RenameError(f"Output mask uses unknown field: {name}")

        value = values[name]
        if isinstance(value, datetime):
            return value.strftime(to_strftime_format(fmt or fields[name].format or "YYYYMMDD"))
        if fmt:
            raise RenameError(f"Field is not a date and cannot be formatted: {name}")
        return str(value)

    rendered = PLACEHOLDER_RE.sub(replace, mask)
    if PLACEHOLDER_RE.search(rendered):
        raise RenameError(f"Could not render output mask: {mask}")
    return rendered


def validate_plan(plans: list[RenamePlan], *, overwrite: bool) -> None:
    targets_seen: dict[Path, Path] = {}
    sources = {plan.source.resolve() for plan in plans}

    for plan in plans:
        target_key = plan.target.resolve()
        if target_key in targets_seen:
            raise RenameError(
                f"Target name conflict: {targets_seen[target_key].name} and {plan.source.name} -> {plan.target.name}"
            )
        targets_seen[target_key] = plan.source

        if plan.target.exists() and target_key not in sources and not overwrite:
            raise RenameError(f"Target already exists: {plan.target}")


def regex_for_spec(spec: FieldSpec) -> str:
    if spec.format and is_date_format(spec.format):
        return regex_for_date_format(spec.format)
    return ".+?"


def is_date_format(fmt: str) -> bool:
    return any(token in fmt for token in DATE_TOKENS)


def regex_for_date_format(fmt: str) -> str:
    pattern = re.escape(fmt)
    replacements = {
        "YYYY": r"\d{4}",
        "YY": r"\d{2}",
        "MM": r"\d{2}",
        "DD": r"\d{2}",
    }
    for token, token_pattern in replacements.items():
        pattern = pattern.replace(re.escape(token), token_pattern)
    return pattern


def parse_date(value: str, fmt: str) -> datetime:
    try:
        return datetime.strptime(value, to_strftime_format(fmt))
    except ValueError as exc:
        raise RenameError(f"Date '{value}' does not match format '{fmt}'") from exc


def to_strftime_format(fmt: str) -> str:
    converted = fmt
    replacements = {
        "YYYY": "%Y",
        "YY": "%y",
        "MM": "%m",
        "DD": "%d",
    }
    for token, strftime_token in replacements.items():
        converted = converted.replace(token, strftime_token)
    return converted
