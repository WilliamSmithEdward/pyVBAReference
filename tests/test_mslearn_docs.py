"""Tests for the Microsoft Learn text cleaner.

``_clean`` runs over every summary, parameter description, remark and constant
description that reaches the generated reference, so its edge cases are the VBA
identifiers that look like markdown: Worksheet_Change, xlPart_, _CodeName.
"""

from __future__ import annotations

import mslearn_docs as docs


def test_asterisk_emphasis_is_unwrapped():
    assert docs._clean("a *star* here") == "a star here"
    assert docs._clean("a **bold** here") == "a bold here"


def test_underscore_emphasis_is_unwrapped():
    assert docs._clean("where _index_ is the name") == "where index is the name"
    assert docs._clean("a __bold__ word") == "a bold word"
    assert docs._clean("_Leading_ and _trailing_") == "Leading and trailing"


def test_the_reported_case():
    raw = ("Use AddIns (_index_), where _index_ is the add-in name or index "
           "number.")
    assert docs._clean(raw) == (
        "Use AddIns (index), where index is the add-in name or index number.")


def test_identifiers_keep_their_underscores():
    # Emphasis cannot start or end inside a word, so none of these are touched.
    for identifier in ("Worksheet_Change", "xlPart_", "_CodeName",
                       "Private_Sub_Name", "a snake_case_name here",
                       "the _Application property"):
        assert docs._clean(identifier) == identifier


def test_emphasised_identifier_keeps_its_inner_underscores():
    # _Worksheet_Change_ is one emphasised run, not _Worksheet_ + Change_.
    assert docs._clean("the _Worksheet_Change_ event") == (
        "the Worksheet_Change event")


def test_underscore_run_must_not_be_padded_with_spaces():
    assert docs._clean("2 _ 3 _ 4") == "2 _ 3 _ 4"


def test_links_are_reduced_to_their_text():
    assert docs._clean("see [Range](excel.range(object).md) now") == (
        "see Range now")


def test_backticks_breaks_and_whitespace():
    assert docs._clean("a `code` word") == "a code word"
    assert docs._clean("one<br>two<BR />three") == "one two three"
    assert docs._clean("  spread   out \n across lines  ") == (
        "spread out across lines")


def test_output_is_ascii():
    assert docs._clean("“quoted” – dashed…") == (
        '"quoted" - dashed...')


def test_strip_name_still_reduces_table_cells():
    assert docs._strip_name("_Password_") == "Password"
    assert docs._strip_name("**xlCSV**") == "xlCSV"
