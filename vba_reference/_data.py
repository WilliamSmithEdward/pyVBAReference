"""Locates and loads the generated JSON reference data.

The data is resolved in two ways:

1. **Installed**: bundled under ``vba_reference/data`` (added at build time).
2. **Development**: the repository root that contains ``index.json`` (the layout
   produced directly by ``scrape_excel_object_model.py``).

Both layouts share the same internal structure::

    <root>/index.json
    <root>/members.json
    <root>/<library>/json/<Type>.json
"""

from __future__ import annotations

import json
import re
from functools import lru_cache
from importlib import resources
from pathlib import Path
from typing import Optional

# Mirrors ``_INVALID`` / ``safe_filename`` in scrape_excel_object_model.py so
# type names map to the same file names.
_INVALID = re.compile(r'[<>:"/\\|?*]')


def safe_filename(name: str) -> str:
    """Return the on-disk file stem used for a type ``name``."""
    return _INVALID.sub("_", name)


class DataNotFoundError(FileNotFoundError):
    """Raised when the generated reference data cannot be located."""


@lru_cache(maxsize=1)
def data_root() -> Path:
    """Return the directory that holds ``index.json`` and the library folders."""
    # 1) Bundled package data (installed wheel).
    try:
        bundled = Path(str(resources.files("vba_reference"))) / "data"
        if (bundled / "index.json").is_file():
            return bundled
    except (ModuleNotFoundError, AttributeError):
        pass
    # 2) Development: walk up from this file looking for the generated data,
    #    which lives in the repo's ``reference/`` subfolder.
    here = Path(__file__).resolve()
    for parent in here.parents:
        for candidate in (parent / "reference", parent):
            if (candidate / "index.json").is_file():
                return candidate
    raise DataNotFoundError(
        "VBA reference data not found. If running from source, generate it "
        "with `python scrape_excel_object_model.py` first."
    )


@lru_cache(maxsize=1)
def master_index() -> dict:
    """The parsed repo-root ``index.json`` (libraries and their types)."""
    return json.loads((data_root() / "index.json").read_text("utf-8"))


@lru_cache(maxsize=1)
def members_index() -> dict:
    """The parsed repo-root ``members.json`` (member name -> definitions)."""
    return json.loads((data_root() / "members.json").read_text("utf-8"))


@lru_cache(maxsize=None)
def load_type_json(library: str, type_name: str) -> dict:
    """Load and parse the raw JSON for one type in ``library``."""
    path = data_root() / library / "json" / f"{safe_filename(type_name)}.json"
    if not path.is_file():
        raise KeyError(f"No type '{type_name}' in library '{library}'.")
    return json.loads(path.read_text("utf-8"))


@lru_cache(maxsize=1)
def type_locations() -> dict[str, list[tuple[str, str, str]]]:
    """Map a lowercased type name to ``[(library_folder, canonical_name, kind)]``."""
    out: dict[str, list[tuple[str, str, str]]] = {}
    for lib in master_index()["libraries"]:
        folder = lib["folder"]
        for t in lib["types"]:
            out.setdefault(t["name"].lower(), []).append(
                (folder, t["name"], t["kind"])
            )
    return out


@lru_cache(maxsize=1)
def members_lower() -> dict[str, tuple[str, list]]:
    """Case-insensitive view of ``members.json``: lower name -> (exact, defs)."""
    out: dict[str, tuple[str, list]] = {}
    for name, defs in members_index()["members"].items():
        out.setdefault(name.lower(), (name, defs))
    return out
