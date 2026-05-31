"""Scrape the VBA-referenceable object models from the live (registered) type
libraries and write one Markdown reference file per class.

Mirrors what the VBA Object Browser shows for each library: classes
(coclasses / dispatch interfaces), their properties / methods / events with
return types and parameter signatures, plus enumerations and their constants.

Libraries scraped (each into its own folder):
    * Excel  -> excel/   (Microsoft Excel Object Library)
    * Office -> office/   (Microsoft Office Object Library)
    * VBA    -> vba/      (Visual Basic For Applications)
    * stdole -> stdole/   (OLE Automation)

VBAProject (the open workbook's own project) is intentionally skipped: it is
not a reusable reference library and is not registered for discovery.

Requirements:
    * Windows with Microsoft Office installed
    * pywin32  (pip install pywin32)

Usage:
    python scrape_excel_object_model.py
"""

from __future__ import annotations

import os
import re
import sys

import pythoncom
from win32com.client import selecttlb


# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))

# Each library: (output folder, include substrings, exclude substrings).
# A type library matches when every "include" substring is present in its
# description and no "exclude" substring is present (case-insensitive).
LIBRARIES = [
    # --- Default references in every Excel VBA project ---
    ("excel", ["microsoft excel", "object library"], []),
    ("office", ["microsoft office", "object library"],
     ["access database engine"]),
    ("vba", ["visual basic for applications"], ["extensibility"]),
    ("stdole", ["ole automation"], []),
    # --- Common references users add via Tools > References ---
    ("scripting", ["microsoft scripting runtime"], []),
    ("msforms", ["microsoft forms 2.0 object library"], []),
    ("adodb", ["microsoft activex data objects"],
     ["recordset", "multi-dimensional", "ext."]),
    ("vbide", ["visual basic for applications extensibility"], []),
    ("msxml", ["microsoft xml"], []),
    ("winhttp", ["winhttp"], []),
]


# Friendly VBA-style names for the common VARIANT types.
VT_NAMES = {
    pythoncom.VT_EMPTY: "",
    pythoncom.VT_NULL: "Null",
    pythoncom.VT_I2: "Integer",
    pythoncom.VT_I4: "Long",
    pythoncom.VT_R4: "Single",
    pythoncom.VT_R8: "Double",
    pythoncom.VT_CY: "Currency",
    pythoncom.VT_DATE: "Date",
    pythoncom.VT_BSTR: "String",
    pythoncom.VT_DISPATCH: "Object",
    pythoncom.VT_ERROR: "Error",
    pythoncom.VT_BOOL: "Boolean",
    pythoncom.VT_VARIANT: "Variant",
    pythoncom.VT_UNKNOWN: "IUnknown",
    pythoncom.VT_DECIMAL: "Decimal",
    pythoncom.VT_I1: "Byte",
    pythoncom.VT_UI1: "Byte",
    pythoncom.VT_UI2: "Integer",
    pythoncom.VT_UI4: "Long",
    pythoncom.VT_I8: "LongLong",
    pythoncom.VT_UI8: "LongLong",
    pythoncom.VT_INT: "Long",
    pythoncom.VT_UINT: "Long",
    pythoncom.VT_VOID: "void",
    pythoncom.VT_HRESULT: "HRESULT",
    pythoncom.VT_LPSTR: "String",
    pythoncom.VT_LPWSTR: "String",
}

# TYPEKIND values
TKIND_ENUM = 0
TKIND_RECORD = 1
TKIND_MODULE = 2
TKIND_INTERFACE = 3
TKIND_DISPATCH = 4
TKIND_COCLASS = 5
TKIND_ALIAS = 6
TKIND_UNION = 7

TKIND_LABEL = {
    TKIND_ENUM: "Enumeration",
    TKIND_RECORD: "Record / UDT",
    TKIND_MODULE: "Module",
    TKIND_INTERFACE: "Interface",
    TKIND_DISPATCH: "Dispatch Interface",
    TKIND_COCLASS: "Class",
    TKIND_ALIAS: "Alias",
    TKIND_UNION: "Union",
}

# Flag bits (defined here so the script does not depend on their presence
# as named attributes across pywin32 versions).
FUNCFLAG_FRESTRICTED = 0x1
FUNCFLAG_FHIDDEN = 0x40
VARFLAG_FREADONLY = 0x1
VARFLAG_FHIDDEN = 0x40
VARFLAG_FRESTRICTED = 0x80
PARAMFLAG_FOPT = 0x10
IMPLTYPEFLAG_FDEFAULT = 0x1
IMPLTYPEFLAG_FSOURCE = 0x2


# --------------------------------------------------------------------------- #
# Type-library loading
# --------------------------------------------------------------------------- #

def load_typelib(include, exclude):
    """Locate and load the highest-version registered type library whose
    description matches the include/exclude substring filters."""
    best = None
    for spec in selecttlb.EnumTlbs():
        desc = (spec.desc or "").lower()
        if not all(sub in desc for sub in include):
            continue
        if any(sub in desc for sub in exclude):
            continue
        key = (_to_int(spec.major), _to_int(spec.minor))
        if best is None or key > best[0]:
            best = (key, spec)

    if best is None:
        return None, None

    spec = best[1]
    try:
        tlb = pythoncom.LoadRegTypeLib(
            spec.clsid, _to_int(spec.major), _to_int(spec.minor),
            _to_int(spec.lcid),
        )
    except pythoncom.com_error:
        # Fall back to loading directly from the on-disk file.
        tlb = pythoncom.LoadTypeLib(spec.dll)
    return tlb, spec.desc


def _to_int(value) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        try:
            return int(str(value), 16)
        except (TypeError, ValueError):
            return 0


# --------------------------------------------------------------------------- #
# Type resolution
# --------------------------------------------------------------------------- #

def resolve_type(tdesc, info) -> str:
    """Turn a TYPEDESC representation into a readable type name."""
    if tdesc is None:
        return ""
    if isinstance(tdesc, int):
        return VT_NAMES.get(tdesc, f"VT_{tdesc}")

    vt = tdesc[0]
    if vt == pythoncom.VT_PTR:
        # Dereference pointers for readability (ByRef params, object refs).
        return resolve_type(tdesc[1], info)
    if vt == pythoncom.VT_SAFEARRAY:
        return f"SAFEARRAY({resolve_type(tdesc[1], info)})"
    if vt == pythoncom.VT_CARRAY:
        return f"{resolve_type(tdesc[1], info)}()"
    if vt == pythoncom.VT_USERDEFINED:
        try:
            ref = info.GetRefTypeInfo(tdesc[1])
            return ref.GetDocumentation(-1)[0]
        except pythoncom.com_error:
            return "Unknown"
    return VT_NAMES.get(vt, f"VT_{vt}")


# --------------------------------------------------------------------------- #
# Member extraction
# --------------------------------------------------------------------------- #

class Member:
    def __init__(self, name: str, kind: str, text: str, doc: str = ""):
        self.name = name
        self.kind = kind  # property | method | event
        self.text = text
        self.doc = doc


def _is_hidden_func(flags: int) -> bool:
    return bool(flags & (FUNCFLAG_FHIDDEN | FUNCFLAG_FRESTRICTED))


def _is_hidden_var(flags: int) -> bool:
    return bool(flags & (VARFLAG_FHIDDEN | VARFLAG_FRESTRICTED))


def _ret_typedesc(fd):
    """Return the TYPEDESC of a function's return value.

    pywin32 exposes ``rettype`` as an ELEMDESC ``(typedesc, paramdesc)``;
    the actual type description is the first element.
    """
    rt = getattr(fd, "rettype", None)
    if rt is None:
        return None
    try:
        return rt[0]
    except (TypeError, IndexError):
        return rt


def _build_signature(info, fd, names) -> str:
    fname = names[0]
    argnames = list(names[1:])
    args = getattr(fd, "args", None) or ()
    parts = []
    for idx in range(len(args)):
        aname = argnames[idx] if idx < len(argnames) else f"Arg{idx + 1}"
        tdesc = None
        pflags = 0
        try:
            elem = args[idx]
            tdesc = elem[0]
            if len(elem) > 1 and isinstance(elem[1], int):
                pflags = elem[1]
        except (IndexError, TypeError):
            pass
        atype = resolve_type(tdesc, info)
        part = aname
        if atype:
            part += f" As {atype}"
        if pflags & PARAMFLAG_FOPT:
            part = f"[{part}]"
        parts.append(part)

    rettype = resolve_type(_ret_typedesc(fd), info)
    sig = f"{fname}({', '.join(parts)})"
    if rettype and rettype not in ("void", "HRESULT", "Null"):
        sig += f" As {rettype}"
    return sig


def extract_interface_members(info):
    """Return (properties, methods) lists for a dispatch/interface typeinfo."""
    attr = info.GetTypeAttr()
    prop_access: dict[str, set] = {}
    prop_type: dict[str, str] = {}
    prop_doc: dict[str, str] = {}
    prop_order: list[str] = []
    methods: list[Member] = []

    for i in range(attr.cFuncs):
        try:
            fd = info.GetFuncDesc(i)
        except pythoncom.com_error:
            continue
        if _is_hidden_func(getattr(fd, "wFuncFlags", 0)):
            continue
        names = info.GetNames(fd.memid)
        if not names:
            continue
        fname = names[0]
        doc = info.GetDocumentation(fd.memid)[1] or ""
        invkind = fd.invkind

        if invkind == pythoncom.INVOKE_PROPERTYGET:
            prop_access.setdefault(fname, set()).add("get")
            if fname not in prop_order:
                prop_order.append(fname)
            prop_type[fname] = resolve_type(_ret_typedesc(fd), info)
            if doc and fname not in prop_doc:
                prop_doc[fname] = doc
        elif invkind in (pythoncom.INVOKE_PROPERTYPUT,
                         pythoncom.INVOKE_PROPERTYPUTREF):
            prop_access.setdefault(fname, set()).add("set")
            if fname not in prop_order:
                prop_order.append(fname)
            if doc and fname not in prop_doc:
                prop_doc[fname] = doc
            # The property's type is the (last) parameter for setters.
            if fname not in prop_type and getattr(fd, "args", None):
                try:
                    prop_type[fname] = resolve_type(fd.args[-1][0], info)
                except (IndexError, TypeError):
                    pass
        else:
            methods.append(Member(fname, "method",
                                  _build_signature(info, fd, names), doc))

    # Dispatch properties expressed as variables.
    for i in range(attr.cVars):
        try:
            vd = info.GetVarDesc(i)
        except pythoncom.com_error:
            continue
        if _is_hidden_var(getattr(vd, "wVarFlags", 0)):
            continue
        names = info.GetNames(vd.memid)
        if not names:
            continue
        vname = names[0]
        prop_access.setdefault(vname, set()).add("get")
        if getattr(vd, "wVarFlags", 0) & VARFLAG_FREADONLY == 0:
            prop_access[vname].add("set")
        if vname not in prop_order:
            prop_order.append(vname)
        try:
            prop_type[vname] = resolve_type(vd.elemdescVar[0], info)
        except (AttributeError, IndexError, TypeError):
            pass

    properties = []
    for name in prop_order:
        access = prop_access.get(name, set())
        acc = "read/write" if {"get", "set"} <= access else (
            "read-only" if "get" in access else "write-only")
        ptype = prop_type.get(name, "")
        text = name + (f" As {ptype}" if ptype else "") + f"  ({acc})"
        properties.append(Member(name, "property", text,
                                 prop_doc.get(name, "")))

    return properties, methods


def extract_events(info):
    attr = info.GetTypeAttr()
    events: list[Member] = []
    for i in range(attr.cFuncs):
        try:
            fd = info.GetFuncDesc(i)
        except pythoncom.com_error:
            continue
        if _is_hidden_func(getattr(fd, "wFuncFlags", 0)):
            continue
        names = info.GetNames(fd.memid)
        if not names:
            continue
        doc = info.GetDocumentation(fd.memid)[1] or ""
        events.append(Member(names[0], "event",
                             _build_signature(info, fd, names), doc))
    return events


def extract_enum_constants(info):
    attr = info.GetTypeAttr()
    consts = []
    for i in range(attr.cVars):
        try:
            vd = info.GetVarDesc(i)
        except pythoncom.com_error:
            continue
        if _is_hidden_var(getattr(vd, "wVarFlags", 0)):
            continue
        names = info.GetNames(vd.memid)
        if not names:
            continue
        consts.append((names[0], getattr(vd, "value", None)))
    return consts


def extract_module_members(info):
    """Return (functions, constants) for a TKIND_MODULE typeinfo.

    Modules hold the global VBA routines (MsgBox, Format, CreateObject, ...)
    and intrinsic constants (vbCrLf, vbTab, ...).
    """
    attr = info.GetTypeAttr()
    funcs: list[Member] = []
    consts: list[tuple] = []
    for i in range(attr.cFuncs):
        try:
            fd = info.GetFuncDesc(i)
        except pythoncom.com_error:
            continue
        if _is_hidden_func(getattr(fd, "wFuncFlags", 0)):
            continue
        names = info.GetNames(fd.memid)
        if not names:
            continue
        doc = info.GetDocumentation(fd.memid)[1] or ""
        funcs.append(Member(names[0], "function",
                            _build_signature(info, fd, names), doc))
    for i in range(attr.cVars):
        try:
            vd = info.GetVarDesc(i)
        except pythoncom.com_error:
            continue
        if _is_hidden_var(getattr(vd, "wVarFlags", 0)):
            continue
        names = info.GetNames(vd.memid)
        if not names:
            continue
        ctype = ""
        try:
            ctype = resolve_type(vd.elemdescVar[0], info)
        except (AttributeError, IndexError, TypeError):
            pass
        consts.append((names[0], getattr(vd, "value", None), ctype))
    return funcs, consts


# --------------------------------------------------------------------------- #
# Coclass resolution
# --------------------------------------------------------------------------- #

def resolve_coclass_interfaces(info):
    """Return (default_interface, source_interface) typeinfos for a coclass."""
    attr = info.GetTypeAttr()
    default_iface = None
    source_iface = None
    for j in range(attr.cImplTypes):
        try:
            flags = info.GetImplTypeFlags(j)
            ref = info.GetRefTypeInfo(info.GetRefTypeOfImplType(j))
        except pythoncom.com_error:
            continue
        if flags & IMPLTYPEFLAG_FSOURCE:
            if source_iface is None or flags & IMPLTYPEFLAG_FDEFAULT:
                source_iface = ref
        else:
            if default_iface is None or flags & IMPLTYPEFLAG_FDEFAULT:
                default_iface = ref
    return default_iface, source_iface


# --------------------------------------------------------------------------- #
# Markdown writing
# --------------------------------------------------------------------------- #

_INVALID = re.compile(r'[<>:"/\\|?*]')


def safe_filename(name: str) -> str:
    return _INVALID.sub("_", name)


def write_class_file(output_dir: str, name: str, kind_label: str, guid: str,
                     lib_desc: str, properties, methods, events,
                     description: str = "") -> None:
    lines = [f"# {name}", ""]
    lines.append(f"**Type:** {kind_label}  ")
    lines.append(f"**Library:** {lib_desc}  ")
    if guid:
        lines.append(f"**GUID:** {guid}  ")
    lines.append("")
    if description:
        lines.append(description)
        lines.append("")

    def emit(members):
        for m in members:
            line = f"- `{m.text}`"
            if m.doc:
                line += f"  \n  {m.doc}"
            lines.append(line)

    if properties:
        lines.append(f"## Properties ({len(properties)})")
        lines.append("")
        emit(properties)
        lines.append("")

    if methods:
        lines.append(f"## Methods ({len(methods)})")
        lines.append("")
        emit(methods)
        lines.append("")

    if events:
        lines.append(f"## Events ({len(events)})")
        lines.append("")
        emit(events)
        lines.append("")

    if not (properties or methods or events):
        lines.append("_No public members._")
        lines.append("")

    path = os.path.join(output_dir, safe_filename(name) + ".md")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines))


def write_enum_file(output_dir: str, name: str, lib_desc: str, consts) -> None:
    lines = [f"# {name}", "", "**Type:** Enumeration  ",
             f"**Library:** {lib_desc}  ", ""]
    lines.append(f"## Constants ({len(consts)})")
    lines.append("")
    for cname, value in consts:
        if value is None:
            lines.append(f"- `{cname}`")
        else:
            lines.append(f"- `{cname}` = {value}")
    lines.append("")
    path = os.path.join(output_dir, safe_filename(name) + ".md")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines))


def write_module_file(output_dir: str, name: str, lib_desc: str, funcs,
                      consts, description: str = "") -> None:
    lines = [f"# {name}", "", "**Type:** Module  ",
             f"**Library:** {lib_desc}  ", ""]
    if description:
        lines.append(description)
        lines.append("")

    if consts:
        lines.append(f"## Constants ({len(consts)})")
        lines.append("")
        for cname, value, ctype in consts:
            t = f" As {ctype}" if ctype else ""
            v = f" = {value}" if value is not None else ""
            lines.append(f"- `{cname}{t}{v}`")
        lines.append("")

    if funcs:
        lines.append(f"## Functions ({len(funcs)})")
        lines.append("")
        for m in funcs:
            line = f"- `{m.text}`"
            if m.doc:
                line += f"  \n  {m.doc}"
            lines.append(line)
        lines.append("")

    path = os.path.join(output_dir, safe_filename(name) + ".md")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines))

def scrape_typelib(tlb, lib_desc: str, output_dir: str) -> int:
    """Scrape one type library into ``output_dir``. Returns files written."""
    os.makedirs(output_dir, exist_ok=True)
    count = tlb.GetTypeInfoCount()

    emitted: set[str] = set()
    written = 0
    index_entries: list[tuple[str, str]] = []

    def display_name(raw: str) -> str:
        # The Object Browser shows coclass names; the underlying dispatch
        # interfaces are typically prefixed with a single underscore.
        return raw[1:] if raw.startswith("_") and len(raw) > 1 else raw

    # Pass 1: coclasses -> rich Application/Range/Worksheet style files.
    for i in range(count):
        try:
            info = tlb.GetTypeInfo(i)
            attr = info.GetTypeAttr()
        except pythoncom.com_error:
            continue
        if attr.typekind != TKIND_COCLASS:
            continue
        name = tlb.GetDocumentation(i)[0]
        description = tlb.GetDocumentation(i)[1] or ""
        default_iface, source_iface = resolve_coclass_interfaces(info)
        props, methods = ([], [])
        events = []
        if default_iface is not None:
            props, methods = extract_interface_members(default_iface)
        if source_iface is not None:
            events = extract_events(source_iface)
        guid = str(attr.iid)
        write_class_file(output_dir, name, TKIND_LABEL[TKIND_COCLASS], guid,
                         lib_desc, props, methods, events, description)
        emitted.add(name)
        index_entries.append((name, "Class"))
        written += 1

    # Pass 2: enumerations.
    for i in range(count):
        try:
            info = tlb.GetTypeInfo(i)
            attr = info.GetTypeAttr()
        except pythoncom.com_error:
            continue
        if attr.typekind != TKIND_ENUM:
            continue
        name = tlb.GetDocumentation(i)[0]
        if name in emitted:
            continue
        consts = extract_enum_constants(info)
        if not consts:
            continue
        write_enum_file(output_dir, name, lib_desc, consts)
        emitted.add(name)
        index_entries.append((name, "Enumeration"))
        written += 1

    # Pass 2b: modules (global functions like MsgBox / Format, and intrinsic
    # constants like vbCrLf).
    for i in range(count):
        try:
            info = tlb.GetTypeInfo(i)
            attr = info.GetTypeAttr()
        except pythoncom.com_error:
            continue
        if attr.typekind != TKIND_MODULE:
            continue
        name = tlb.GetDocumentation(i)[0]
        if name in emitted:
            continue
        description = tlb.GetDocumentation(i)[1] or ""
        funcs, consts = extract_module_members(info)
        if not (funcs or consts):
            continue
        write_module_file(output_dir, name, lib_desc, funcs, consts,
                          description)
        emitted.add(name)
        index_entries.append((name, "Module"))
        written += 1

    # Pass 3: dispatch/interfaces not already represented by a coclass.
    for i in range(count):
        try:
            info = tlb.GetTypeInfo(i)
            attr = info.GetTypeAttr()
        except pythoncom.com_error:
            continue
        if attr.typekind not in (TKIND_DISPATCH, TKIND_INTERFACE):
            continue
        raw = tlb.GetDocumentation(i)[0]
        name = display_name(raw)
        if name in emitted:
            continue
        description = tlb.GetDocumentation(i)[1] or ""
        props, methods = extract_interface_members(info)
        if not (props or methods):
            continue
        write_class_file(output_dir, name, TKIND_LABEL[attr.typekind],
                         str(attr.iid), lib_desc, props, methods, [],
                         description)
        emitted.add(name)
        index_entries.append((name, TKIND_LABEL[attr.typekind]))
        written += 1

    # Index for navigation.
    index_entries.sort(key=lambda e: e[0].lower())
    index_lines = [f"# {lib_desc}", "",
                   f"Scraped object model: {len(index_entries)} entries.", ""]
    for name, kind in index_entries:
        index_lines.append(f"- [{name}]({safe_filename(name)}.md) - {kind}")
    index_lines.append("")
    with open(os.path.join(output_dir, "_index.md"), "w",
              encoding="utf-8") as fh:
        fh.write("\n".join(index_lines))

    return written


def main() -> int:
    total = 0
    libs_done = 0
    for folder, include, exclude in LIBRARIES:
        tlb, desc = load_typelib(include, exclude)
        if tlb is None:
            print(f"[skip] No registered type library matched '{folder}' "
                  f"(include={include}).")
            continue
        lib_doc = tlb.GetDocumentation(-1)
        lib_desc = desc or lib_doc[1] or lib_doc[0] or folder
        output_dir = os.path.join(ROOT_DIR, folder)
        print(f"Scraping {lib_desc} -> {folder}/ "
              f"({tlb.GetTypeInfoCount()} type entries)...")
        written = scrape_typelib(tlb, lib_desc, output_dir)
        print(f"  Wrote {written} files to {folder}/ (index: {folder}/_index.md)")
        total += written
        libs_done += 1

    if libs_done == 0:
        print("No matching type libraries were found. Is Office installed?")
        return 1

    print(f"\nDone. {total} reference files written across "
          f"{libs_done} libraries.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
