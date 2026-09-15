"""Tests for the vba_reference query API (runs against the in-repo data)."""

from __future__ import annotations

import vba_reference as vba


def test_libraries_present():
    names = vba.library_names()
    for host in ("excel", "powerpoint", "word", "access"):
        assert host in names
    assert "vba" in names
    assert len(names) == 13


def test_master_catalog_counts():
    libs = {lib["folder"]: lib for lib in vba.libraries()}
    assert libs["excel"]["type_count"] > 1000
    # type_count matches the listed types.
    assert libs["excel"]["type_count"] == len(libs["excel"]["types"])


def test_get_type_worksheet():
    ws = vba.get_type("Worksheet")
    assert ws.name == "Worksheet"
    assert ws.kind == "Class"
    assert ws.library.startswith("Microsoft Excel")
    assert ws.description
    assert ws.remarks  # enriched from MS Learn


def test_get_type_is_case_insensitive():
    assert vba.get_type("worksheet").name == "Worksheet"


def test_method_parameters_enriched():
    protect = vba.get_member("Worksheet", "Protect")
    assert protect is not None
    assert protect.kind == "method"
    names = [p.name for p in protect.parameters]
    assert "Password" in names
    pwd = next(p for p in protect.parameters if p.name == "Password")
    assert pwd.optional is True
    assert pwd.description  # parameter description present


def test_property_access_mode():
    name = vba.get_member("Worksheet", "Name")
    assert name is not None
    assert name.kind == "property"
    assert name.access in {"read-only", "write-only", "read/write"}


def test_enum_constant_value_and_description():
    csv = vba.get_constant("XlFileFormat", "xlCSV")
    assert csv is not None
    assert csv.value == 6
    assert csv.description


def test_word_document_type():
    doc = vba.get_type("Document", "word")
    assert doc.kind == "Class"
    assert doc.library.startswith("Microsoft Word")
    assert any(p.name == "Paragraphs" for p in doc.properties)


def test_powerpoint_presentation_type():
    pres = vba.get_type("Presentation")
    assert pres.library.startswith("Microsoft PowerPoint")
    assert any(m.name == "SaveAs" for m in pres.methods)


def test_access_docmd_parameters_enriched():
    open_form = vba.get_member("DoCmd", "OpenForm", "access")
    assert open_form is not None
    names = [p.name for p in open_form.parameters]
    assert names[:2] == ["FormName", "View"]
    assert open_form.parameters[0].description


def test_host_application_enum_constants():
    assert vba.get_constant("WdSaveFormat", "wdFormatPDF").value == 17
    assert vba.get_constant("PpSaveAsFileType", "ppSaveAsPDF").value == 32
    assert vba.get_constant("AcFormView", "acNormal").description


def test_shared_type_name_resolves_to_excel_first():
    # Application, Font and Range exist in several hosts; catalog order puts
    # Excel first so the unqualified lookup stays stable.
    assert vba.get_type("Application").library.startswith("Microsoft Excel")
    libs = {r.library for r in vba.locate_type("Application")}
    assert {"excel", "powerpoint", "word", "access"} <= libs


def test_type_names_are_unique_per_library_ignoring_case():
    # A type name is also its file name, so two types in one library whose
    # names differ only in case would overwrite each other on Windows.
    for lib in vba.libraries():
        lowered = [t["name"].lower() for t in lib["types"]]
        assert len(lowered) == len(set(lowered)), lib["folder"]


def test_coclass_wins_over_same_name_interface():
    # Access declares the CheckBox coclass and a Checkbox dispatch interface;
    # the coclass is the richer entry and the one the Object Browser shows.
    check_box = vba.get_type("CheckBox", "access")
    assert check_box.kind == "Class"
    assert check_box.events


def test_find_members_msgbox_is_vba_builtin():
    refs = vba.find_members("MsgBox")
    assert any(r.library == "vba" and r.type == "Interaction"
               and r.kind == "function" for r in refs)


def test_find_members_case_insensitive():
    assert vba.find_members("msgbox")  # lower-case still resolves


def test_find_members_saveas_cross_library():
    refs = vba.find_members("SaveAs")
    types = {r.type for r in refs}
    assert "Workbook" in types


def test_locate_type_returns_library():
    refs = vba.locate_type("Worksheet")
    assert any(r.library == "excel" for r in refs)


def test_search_types_substring():
    hits = vba.search_types("worksheet")
    assert any(r.name == "Worksheet" for r in hits)


def test_unknown_type_raises():
    import pytest

    with pytest.raises(KeyError):
        vba.get_type("NoSuchTypeXYZ")


def test_data_path_resolves():
    import os

    root = vba.data_path()
    assert os.path.isfile(os.path.join(root, "index.json"))
