"""Tests for the consolidated per-application exports.

They are generated from the per-type files, so the checks here are that the
fold-up is complete (no library or type dropped) and faithful (the copied
entries are identical to their source).
"""

from __future__ import annotations

import csv
import json
import os
from functools import lru_cache

import pytest

import vba_reference as vba

HOSTS = ("excel", "powerpoint", "word", "access")
SHARED = "shared"
KEYS = HOSTS + (SHARED,)
SUFFIXES = ("", "_constants", "_properties")
CSV_HEADER = ["Object", "Property", "Property split", "Type", "Access"]
CONSTANTS_CSV_HEADER = ["Owner", "Kind", "Constant", "Value", "Description"]
ACCESS_CODES = {"R_O", "V", "W_O"}


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


@lru_cache(maxsize=None)
def rows(name: str) -> tuple:
    with open(os.path.join(consolidated_dir(), name + ".csv"),
              encoding="utf-8", newline="") as fh:
        return tuple(tuple(row) for row in csv.reader(fh))


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


def test_property_csv_exists_for_every_application():
    for key in KEYS:
        table = rows(key + "_properties")
        assert list(table[0]) == CSV_HEADER
        assert len(table) > 1


def test_property_csv_matches_the_properties_export():
    for key in KEYS:
        table = rows(key + "_properties")
        doc = export(key + "_properties")
        assert len(table) - 1 == doc["property_count"]
        from_json = [(t["name"], p["name"])
                     for lib in doc["libraries"] for t in lib["types"]
                     for p in t["properties"]]
        assert [(r[0], r[1]) for r in table[1:]] == from_json


def test_property_csv_row_shape():
    # The row the request asked for: Application | ActiveCell | Range | R_O,
    # with the name also split at capitals.
    table = rows("excel_properties")
    active_cell = next(r for r in table
                       if r[0] == "Application" and r[1] == "ActiveCell")
    assert active_cell == ("Application", "ActiveCell", "Active Cell",
                           "Range", "R_O")
    used_range = next(r for r in table
                      if r[0] == "Worksheet" and r[1] == "UsedRange")
    assert used_range[2] == "Used Range"


def test_property_csv_access_codes():
    seen = set()
    for key in KEYS:
        for row in rows(key + "_properties")[1:]:
            assert row[4] in ACCESS_CODES, row
            seen.add(row[4])
    assert {"R_O", "V"} <= seen


def test_property_csv_is_safe_to_open_in_a_spreadsheet():
    # A leading =, +, - or @ makes a spreadsheet evaluate the cell as a
    # formula. Identifiers never start that way; assert it rather than assume.
    for key in KEYS:
        for row in rows(key + "_properties")[1:]:
            for field in row:
                assert field[:1] not in ("=", "+", "-", "@"), row


def test_constants_csv_exists_for_every_application():
    for key in KEYS:
        table = rows(key + "_constants")
        assert list(table[0]) == CONSTANTS_CSV_HEADER
        assert len(table) > 1


def test_constants_csv_matches_the_constants_export():
    for key in KEYS:
        table = rows(key + "_constants")
        doc = export(key + "_constants")
        assert len(table) - 1 == doc["constant_count"]
        from_json = [(e["name"], c["name"])
                     for lib in doc["libraries"] for e in lib["enumerations"]
                     for c in e["constants"]]
        assert [(r[0], r[2]) for r in table[1:]] == from_json


def test_constants_csv_row_shape():
    table = rows("excel_constants")
    csv_format = next(r for r in table
                      if r[0] == "XlFileFormat" and r[2] == "xlCSV")
    assert csv_format == ("XlFileFormat", "Enumeration", "xlCSV", "6", "CSV")
    # Negative values stay usable as numbers.
    assert any(r[3] == "-4104" for r in table[1:])
    assert {r[1] for r in table[1:]} <= {"Enumeration", "Module"}


def test_constants_csv_renders_values_as_vb_literals():
    table = rows("shared_constants")
    by_name = {r[2]: r[3] for r in table[1:]}
    # A constant whose value *is* a control character must not put raw bytes
    # in the file.
    assert by_name["vbCrLf"] == "Chr(13) & Chr(10)"
    assert by_name["vbTab"] == "Chr(9)"
    assert by_name["vbNullChar"] == "Chr(0)"
    assert by_name["vbNullString"] == '""'
    quoted = rows("access_constants")
    assert any(r[3] == '"PDF Format (*.pdf)"' for r in quoted[1:])


def test_constants_csv_holds_no_control_characters():
    for key in KEYS:
        for row in rows(key + "_constants")[1:]:
            for field in row:
                assert all(ch == " " or ch.isprintable() for ch in field), row


def test_constants_csv_is_safe_to_open_in_a_spreadsheet():
    # As for the property table, but enum values are legitimately negative,
    # so a leading "-" is only allowed when the field really is a number.
    for key in KEYS:
        for row in rows(key + "_constants")[1:]:
            for field in row:
                assert field[:1] not in ("=", "+", "@"), row
                if field[:1] == "-":
                    int(field)


def test_formula_leading_text_is_kept_as_text():
    # wdFieldExpression is described as "= (Formula) field.", which a
    # spreadsheet would evaluate. It survives with a leading space.
    table = rows("word_constants")
    expression = next(r for r in table[1:] if r[2] == "wdFieldExpression")
    assert expression[4] == " = (Formula) field."
    assert expression[4].strip() == "= (Formula) field."


def test_markdown_covers_every_type():
    for key in HOSTS:
        text = markdown(key)
        for entry in vba.list_types(key):
            assert f"\n### {entry['name']}\n" in text, f"{key}/{entry['name']}"
