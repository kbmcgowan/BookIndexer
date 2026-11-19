"""Shared TOML preferences for the index candidate tools."""

import json
import os
import tempfile
import tomllib
from dataclasses import dataclass
from pathlib import Path

PREFERENCES_PATH = Path(__file__).resolve().with_name("preferences.toml")
DEFAULT_INDEX_PATH = "./index/"
DEFAULT_BOOK_PATH = "./book/"
HIDEABLE_FIELDS = (
    "source",
    "document",
    "tfidf_score",
    "occurrences",
    "rank",
    "entity_type",
)


@dataclass(frozen=True)
class Preferences:
    index_path: Path
    book_path: Path
    hidden_fields: tuple[str, ...]


def resolve_preference_path(
    value: str | Path,
    config_path: Path = PREFERENCES_PATH,
) -> Path:
    path = Path(value).expanduser()
    if not path.is_absolute():
        path = config_path.resolve().parent / path
    return path.resolve()


def load_preferences(config_path: Path = PREFERENCES_PATH) -> Preferences:
    config_path = config_path.expanduser().resolve()
    if config_path.is_file():
        with config_path.open("rb") as config_file:
            values = tomllib.load(config_file)
    else:
        values = {}

    index_path = values.get("index_path", DEFAULT_INDEX_PATH)
    book_path = values.get("book_path", DEFAULT_BOOK_PATH)
    hidden_fields = values.get("hidden_fields", ["source"])

    if not isinstance(index_path, str) or not index_path.strip():
        raise ValueError("'index_path' must be a non-empty path string.")
    if not isinstance(book_path, str) or not book_path.strip():
        raise ValueError("'book_path' must be a non-empty path string.")
    if not isinstance(hidden_fields, list) or not all(
        isinstance(field, str) for field in hidden_fields
    ):
        raise ValueError("'hidden_fields' must be a list of field names.")

    unknown_fields = set(hidden_fields) - set(HIDEABLE_FIELDS)
    if unknown_fields:
        raise ValueError(
            "'hidden_fields' contains unsupported field(s): "
            f"{', '.join(sorted(unknown_fields))}."
        )

    return Preferences(
        index_path=resolve_preference_path(index_path, config_path),
        book_path=resolve_preference_path(book_path, config_path),
        hidden_fields=tuple(
            field for field in HIDEABLE_FIELDS if field in set(hidden_fields)
        ),
    )


def _serialize_path(path: Path, config_path: Path) -> str:
    path = path.expanduser().resolve()
    base = config_path.resolve().parent
    try:
        relative_path = path.relative_to(base).as_posix()
    except ValueError:
        return str(path)
    return f"./{relative_path}" if relative_path != "." else "./"


def save_preferences(
    preferences: Preferences,
    config_path: Path = PREFERENCES_PATH,
) -> None:
    config_path = config_path.expanduser().resolve()
    unknown_fields = set(preferences.hidden_fields) - set(HIDEABLE_FIELDS)
    if unknown_fields:
        raise ValueError(
            f"Unsupported hidden field(s): {', '.join(sorted(unknown_fields))}."
        )

    hidden_values = ", ".join(
        json.dumps(field)
        for field in HIDEABLE_FIELDS
        if field in preferences.hidden_fields
    )
    lines = [
        f"index_path = {json.dumps(_serialize_path(preferences.index_path, config_path))}",
        f"book_path = {json.dumps(_serialize_path(preferences.book_path, config_path))}",
        f"hidden_fields = [{hidden_values}]",
    ]

    temporary_path = None
    try:
        config_path.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            dir=config_path.parent,
            prefix=f".{config_path.name}.",
            suffix=".tmp",
            delete=False,
        ) as config_file:
            temporary_path = Path(config_file.name)
            config_file.writelines(f"{line}\n" for line in lines)
        os.replace(temporary_path, config_path)
    finally:
        if temporary_path is not None and temporary_path.exists():
            temporary_path.unlink()
