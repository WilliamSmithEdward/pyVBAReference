"""Tests for the consolidated per-application exports.

They are generated from the per-type files, so the checks here are that the
fold-up is complete (no library or type dropped) and faithful (the copied
entries are identical to their source).
"""

from __future__ import annotations

import json
import os
from functools import lru_cache

import pytest

import vba_reference as vba

HOSTS = ("excel", "powerpoint", "word", "access")
SHARED = "shared"
KEYS = HOSTS + (SHARED,)
SUFFIXES = ("", "_constants", "_properties")


def consolidated_dir() -> str:
    path = os.path.join(vba.data_path(), "consolidated")
    if not os.path.isdir(path):
        pytest.skip("consolidated exports are not part of the installed data")
    return path


@lru_cache(maxsize=None)
def export(name: str) -> dict:
    with open(os.path.join(consolidated_dir(), name + ".json"),
              encoding="utf-8") as fh:
        return json.load(fh)


@lru_cache(maxsize=None)
def markdown(name: str) -> str:
    with open(os.path.join(consolidated_dir(), name + ".md"),
              encoding="utf-8") as fh:
        return fh.read()


def test_every_file_of_every_set_exists():
    for key in KEYS:
        for suffix in SUFFIXES:
            for ext in (".json", ".md"):
                path = os.path.join(consolidated_dir(), key + suffix + ext)
                assert os.path.isfile(path), path


def test_each_host_holds_exactly_its_own_library():
    for key in HOSTS:
        doc = export(key)
        assert [lib["folder"] for lib in doc["libraries"]] == [key]
        assert doc["key"] == key


def test_libraries_are_partitioned_across_the_exports():
    exported = [lib["folder"] for key in KEYS
                for lib in export(key)["libraries"]]
    assert sorted(exported) == sorted(vba.library_names())


def test_types_match_the_catalog_library_by_library():
    total = 0
    for key in KEYS:
        doc = export(key)
        for lib in doc["libraries"]:
            names = [t["name"] for t in lib["types"]]
            assert names == [t["name"] for t in vba.list_types(lib["folder"])]
            assert lib["type_count"] == len(names)
        assert doc["type_count"] == sum(lib["type_count"]
                                        for lib in doc["libraries"])
        total += doc["type_count"]
    assert total == sum(lib["type_count"] for lib in vba.libraries())


def _source_types(key: str) -> dict:
    """Map ``folder -> {type name: entry}`` from the full export of ``key``."""
    return {lib["folder"]: {t["name"]: t for t in lib["types"]}
            for lib in export(key)["libraries"]}


def test_constants_export_is_a_faithful_subset():
    for key in KEYS:
        source = _source_types(key)
        doc = export(key + "_constants")
        counted = 0
        for lib in doc["libraries"]:
            for entry in lib["enumerations"]:
                assert entry["kind"] in ("Enumeration", "Module")
                origin = source[lib["folder"]][entry["name"]]
                assert entry["constants"] == origin["constants"]
                assert entry["constant_count"] == len(entry["constants"])
                counted += entry["constant_count"]
        assert counted == doc["constant_count"]
        assert counted > 0


def test_properties_export_is_a_faithful_subset():
    for key in KEYS:
        source = _source_types(key)
        doc = export(key + "_properties")
        counted = 0
        for lib in doc["libraries"]:
            for entry in lib["types"]:
                origin = source[lib["folder"]][entry["name"]]
                assert entry["properties"] == origin["properties"]
                assert entry["property_count"] == len(entry["properties"])
                counted += entry["property_count"]
        assert counted == doc["property_count"]
        assert counted > 0


def test_excel_constants_carry_values_and_descriptions():
    doc = export("excel_constants")
    file_format = next(e for lib in doc["libraries"]
                       for e in lib["enumerations"]
                       if e["name"] == "XlFileFormat")
    csv = next(c for c in file_format["constants"] if c["name"] == "xlCSV")
    assert csv["value"] == 6
    assert csv["description"]


def test_markdown_headings_and_content():
    for key in KEYS:
        assert markdown(key).startswith("# ")
        assert "\n## " in markdown(key)          # a library section
        assert "\n### " in markdown(key)         # a type
    assert "# Word VBA object model" in markdown("word")
    assert "`Paragraphs As Paragraphs  (read-only)`" in markdown(
        "word_properties")
    assert "`wdFormatPDF` = 17" in markdown("word_constants")


def test_markdown_covers_every_type():
    for key in HOSTS:
        text = markdown(key)
        for entry in vba.list_types(key):
            assert f"\n### {entry['name']}\n" in text, f"{key}/{entry['name']}"
