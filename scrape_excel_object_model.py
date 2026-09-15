"""Scrape the VBA-referenceable object models from the live (registered) type
libraries and write one Markdown reference file per class.

Mirrors what the VBA Object Browser shows for each library: classes
(coclasses / dispatch interfaces), their properties / methods / events with
return types and parameter signatures, plus enumerations and their constants.

Every library listed in ``LIBRARIES`` gets its own folder under reference/:
the four host applications (Excel, PowerPoint, Word, Access) and the shared
libraries every host can reference (Office, VBA, stdole, MSForms, ADO, ...).

VBAProject (a document's own project) is intentionally skipped: it is not a
reusable reference library and is not registered for discovery.

Requirements:
    * Windows with the relevant Microsoft Office applications installed
    * pywin32  (pip install pywin32)

Usage:
    python scrape_excel_object_model.py
    python scrape_excel_object_model.py --only word,powerpoint
"""

from __future__ import annotations

import json
import os
import sys

import pythoncom
from win32com.client import selecttlb

import consolidate_reference
import mslearn_docs
from consolidate_reference import render_type_md, safe_filename


# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))

# Generated reference data lives in its own subfolder to keep the repo root
# uncluttered.
DATA_DIR = os.path.join(ROOT_DIR, "reference")

# Each library: (output folder, include substrings, exclude substrings,
# MS Learn api prefix). The api prefix is the library namespace used by the
# MicrosoftDocs/VBA-Docs reference (e.g. "excel" -> api/Excel.*.md). Use None
# for libraries that have no MS Learn VBA-API coverage; their files are still
# written, just without external descriptions.
LIBRARIES = [
    # --- Host applications ---
    ("excel", ["microsoft excel", "object library"], [], "excel"),
    ("powerpoint", ["microsoft powerpoint", "object library"], [],
     "powerpoint"),
    ("word", ["microsoft word", "object library"], [], "word"),
    ("access", ["microsoft access", "object library"],
     ["database engine"], "access"),
    # --- Shared references available to every host ---
    ("office", ["microsoft office", "object library"],
     ["access database engine"], "office"),
    ("vba", ["visual basic for applications"], ["extensibility"], "vba"),
    ("stdole", ["ole automation"], [], None),
    # --- Common references users add via Tools > References ---
    ("scripting", ["microsoft scripting runtime"], [], None),
    ("msforms", ["microsoft forms 2.0 object library"], [], None),
    ("adodb", ["microsoft activex data objects"],
     ["recordset", "multi-dimensional", "ext."], None),
    ("vbide", ["visual basic for applications extensibility"], [], None),
    ("msxml", ["microsoft xml"], [], None),
    ("winhttp", ["winhttp"], [], None),
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
        self.kind = kind  # property | method | event | function
        self.text = text
        self.doc = doc
        # Structured detail (populated where applicable).
        self.params: list[dict] = []   # [{name, type, optional, description}]
        self.ret: str = ""             # return type
        self.ptype: str = ""           # property type
        self.access: str = ""          # read/write | read-only | write-only


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
    sig, _params = _sig_and_params(info, fd, names)
    return sig


def _sig_and_params(info, fd, names):
    """Return (signature_string, params_list).

    params_list is an ordered list of ``{"name", "type", "optional"}`` dicts;
    parameter descriptions are filled in later from the MS Learn index.
    """
    fname = names[0]
    argnames = list(names[1:])
    args = getattr(fd, "args", None) or ()
    parts = []
    params = []
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
        optional = bool(pflags & PARAMFLAG_FOPT)
        part = aname
        if atype:
            part += f" As {atype}"
        if optional:
            part = f"[{part}]"
        parts.append(part)
        params.append({"name": aname, "type": atype,
                       "optional": optional, "description": ""})

    rettype = resolve_type(_ret_typedesc(fd), info)
    sig = f"{fname}({', '.join(parts)})"
    if rettype and rettype not in ("void", "HRESULT", "Null"):
        sig += f" As {rettype}"
    return sig, params


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
            sig, params = _sig_and_params(info, fd, names)
            mem = Member(fname, "method", sig, doc)
            mem.params = params
            mem.ret = resolve_type(_ret_typedesc(fd), info)
            methods.append(mem)

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
        mem = Member(name, "property", text, prop_doc.get(name, ""))
        mem.ptype = ptype
        mem.access = acc
        properties.append(mem)

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
        sig, params = _sig_and_params(info, fd, names)
        ev = Member(names[0], "event", sig, doc)
        ev.params = params
        events.append(ev)
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
        sig, params = _sig_and_params(info, fd, names)
        fn = Member(names[0], "function", sig, doc)
        fn.params = params
        fn.ret = resolve_type(_ret_typedesc(fd), info)
        funcs.append(fn)
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
# Output writing
# --------------------------------------------------------------------------- #

def _member_json(m) -> dict:
    if m.kind == "property":
        return {"name": m.name, "kind": "property", "type": m.ptype,
                "access": m.access, "description": m.doc}
    d = {"name": m.name, "kind": m.kind, "signature": m.text,
         "description": m.doc}
    if m.ret:
        d["returns"] = m.ret
    if m.params:
        d["parameters"] = [
            {"name": p["name"], "type": p["type"],
             "optional": p["optional"], "description": p["description"]}
            for p in m.params]
    return d


def write_type(md_dir: str, json_dir: str, data: dict) -> None:
    """Write one type as JSON and as Markdown, both from the same dict.

    The Markdown comes from ``consolidate_reference.render_type_md`` so the
    per-type files and the consolidated exports cannot drift apart.
    """
    stem = safe_filename(data["name"])
    with open(os.path.join(json_dir, stem + ".json"), "w",
              encoding="utf-8") as fh:
        json.dump(data, fh, indent=2)
        fh.write("\n")
    with open(os.path.join(md_dir, stem + ".md"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(render_type_md(data)))


def _apply_params(member, entry) -> None:
    """Fill structured parameter descriptions from a DocEntry (by name)."""
    if entry is None or not entry.params:
        return
    for p in member.params:
        desc = entry.params.get(p["name"].lower())
        if desc:
            p["description"] = desc


def _enrich_api(index, prefix, type_name, props, methods, events):
    """Overlay MS Learn docs onto api-backed members. Returns the type's
    DocEntry (for remarks/example) or None."""
    if index is None or not prefix:
        return None
    for m in props:
        e = index.get_member(prefix, type_name, m.name)
        if e and e.summary:
            m.doc = e.summary
    for m in methods:
        e = index.get_member(prefix, type_name, m.name)
        if e:
            if e.summary:
                m.doc = e.summary
            _apply_params(m, e)
    for m in events:
        e = index.get_event(prefix, type_name, m.name)
        if e:
            if e.summary:
                m.doc = e.summary
            _apply_params(m, e)
    return index.get_type(prefix, type_name)


def _enrich_builtins(index, funcs):
    """Overlay VBA language built-in docs (MsgBox, Format, ...) onto module
    functions, matched by name."""
    if index is None:
        return
    for m in funcs:
        e = index.get_builtin(m.name)
        if e:
            if e.summary:
                m.doc = e.summary
            _apply_params(m, e)


def scrape_typelib(tlb, lib_desc: str, output_dir: str, doc_prefix=None,
                   index=None) -> int:
    """Scrape one type library into ``output_dir``.

    Writes one Markdown file per type into ``<output_dir>/md/`` and one JSON
    file per type into ``<output_dir>/json/``. Returns files written.
    """
    md_dir = os.path.join(output_dir, "md")
    json_dir = os.path.join(output_dir, "json")
    os.makedirs(md_dir, exist_ok=True)
    os.makedirs(json_dir, exist_ok=True)
    count = tlb.GetTypeInfoCount()

    # The VBA library matches global built-ins by member name, not api path.
    builtin_mode = (doc_prefix == "vba")
    api_prefix = None if builtin_mode else doc_prefix

    # Lowercased, because a type name is also its file name: Access and VBIDE
    # name a coclass and its dispatch interface identically except for case
    # (CheckBox / Checkbox), and the coclass - emitted first, and the one the
    # Object Browser shows - is the richer entry of the two.
    emitted: set[str] = set()
    written = 0
    index_entries: list[tuple[str, str]] = []

    def display_name(raw: str) -> str:
        # The Object Browser shows coclass names; the underlying dispatch
        # interfaces are typically prefixed with a single underscore.
        return raw[1:] if raw.startswith("_") and len(raw) > 1 else raw

    def emit_class(name, kind_label, guid, props, methods, events,
                   description):
        remarks = example = ""
        if index is not None and api_prefix:
            entry = _enrich_api(index, api_prefix, name, props, methods,
                                events)
            if entry:
                if entry.summary:
                    description = entry.summary
                remarks, example = entry.remarks, entry.example
        write_type(md_dir, json_dir, {
            "name": name, "kind": kind_label, "guid": guid,
            "library": lib_desc, "description": description,
            "remarks": remarks, "example": example,
            "properties": [_member_json(m) for m in props],
            "methods": [_member_json(m) for m in methods],
            "events": [_member_json(m) for m in events],
        })

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
        emit_class(name, TKIND_LABEL[TKIND_COCLASS], str(attr.iid),
                   props, methods, events, description)
        emitted.add(name.lower())
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
        if name.lower() in emitted:
            continue
        raw_consts = extract_enum_constants(info)
        if not raw_consts:
            continue
        description = tlb.GetDocumentation(i)[1] or ""
        const_docs = {}
        if index is not None and api_prefix:
            entry = index.get_type(api_prefix, name)
            if entry:
                if entry.summary:
                    description = entry.summary
                const_docs = entry.constants
        consts = [(cname, value, const_docs.get(cname.lower(), ""))
                  for cname, value in raw_consts]
        write_type(md_dir, json_dir, {
            "name": name, "kind": "Enumeration", "library": lib_desc,
            "description": description,
            "constants": [{"name": c, "value": v, "description": d}
                          for c, v, d in consts],
        })
        emitted.add(name.lower())
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
        if name.lower() in emitted:
            continue
        description = tlb.GetDocumentation(i)[1] or ""
        funcs, consts = extract_module_members(info)
        if not (funcs or consts):
            continue
        if index is not None:
            if builtin_mode:
                _enrich_builtins(index, funcs)
            elif api_prefix:
                entry = _enrich_api(index, api_prefix, name, [], funcs, [])
                if entry and entry.summary:
                    description = entry.summary
        write_type(md_dir, json_dir, {
            "name": name, "kind": "Module", "library": lib_desc,
            "description": description,
            "functions": [_member_json(m) for m in funcs],
            "constants": [{"name": c, "value": v, "type": t}
                          for c, v, t in consts],
        })
        emitted.add(name.lower())
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
        if name.lower() in emitted:
            continue
        description = tlb.GetDocumentation(i)[1] or ""
        props, methods = extract_interface_members(info)
        if not (props or methods):
            continue
        emit_class(name, TKIND_LABEL[attr.typekind], str(attr.iid),
                   props, methods, [], description)
        emitted.add(name.lower())
        index_entries.append((name, TKIND_LABEL[attr.typekind]))
        written += 1

    # Index for navigation (Markdown + JSON).
    index_entries.sort(key=lambda e: e[0].lower())
    index_lines = [f"# {lib_desc}", "",
                   f"Scraped object model: {len(index_entries)} entries.", "",
                   "One Markdown file per type in this folder; matching "
                   "machine-readable JSON in `../json/`.", ""]
    for name, kind in index_entries:
        index_lines.append(f"- [{name}]({safe_filename(name)}.md) - {kind}")
    index_lines.append("")
    with open(os.path.join(md_dir, "_index.md"), "w",
              encoding="utf-8") as fh:
        fh.write("\n".join(index_lines))
    with open(os.path.join(json_dir, "_index.json"), "w",
              encoding="utf-8") as fh:
        json.dump({"library": lib_desc,
                   "types": [{"name": n, "kind": k} for n, k in index_entries]},
                  fh, indent=2)
        fh.write("\n")

    return written


def write_master_indexes(root_dir: str, done_libs) -> int:
    """Build repo-root ``index.json`` and ``members.json`` from the per-type
    JSON already written for each library.

    ``done_libs`` is a list of ``(folder, library_description)`` tuples. The
    master index lists every library and its types; the member index maps each
    member name to the types that define it (for fast cross-library lookup).
    Returns the number of distinct member names indexed.
    """
    libraries = []
    members: dict[str, list] = {}
    total_types = 0
    for folder, lib_desc in done_libs:
        json_dir = os.path.join(root_dir, folder, "json")
        try:
            with open(os.path.join(json_dir, "_index.json"),
                      encoding="utf-8") as fh:
                idx = json.load(fh)
        except OSError:
            continue
        types = idx.get("types", [])
        total_types += len(types)
        libraries.append({
            "folder": folder, "library": lib_desc,
            "type_count": len(types), "types": types,
        })
        for t in types:
            tname = t["name"]
            try:
                with open(os.path.join(json_dir,
                                       f"{safe_filename(tname)}.json"),
                          encoding="utf-8") as fh:
                    data = json.load(fh)
            except OSError:
                continue
            groups = (
                ("property", data.get("properties", [])),
                ("method", data.get("methods", [])),
                ("event", data.get("events", [])),
                ("function", data.get("functions", [])),
                ("constant", data.get("constants", [])),
            )
            for kind, items in groups:
                for m in items:
                    members.setdefault(m["name"], []).append(
                        {"library": folder, "type": tname, "kind": kind})

    members_sorted = {k: members[k] for k in sorted(members, key=str.lower)}
    with open(os.path.join(root_dir, "index.json"), "w",
              encoding="utf-8") as fh:
        json.dump({"total_libraries": len(libraries),
                   "total_types": total_types,
                   "libraries": libraries}, fh, indent=2)
        fh.write("\n")
    with open(os.path.join(root_dir, "members.json"), "w",
              encoding="utf-8") as fh:
        json.dump({"total_names": len(members_sorted),
                   "members": members_sorted}, fh, indent=2)
        fh.write("\n")
    return len(members_sorted)


def _clean_stale_flat(output_dir: str) -> None:
    """Remove files from the previous flat layout (``<folder>/*.md``) so only
    the new ``md/`` and ``json/`` subfolders remain."""
    if not os.path.isdir(output_dir):
        return
    for entry in os.listdir(output_dir):
        full = os.path.join(output_dir, entry)
        if os.path.isfile(full) and entry.lower().endswith((".md", ".json")):
            os.remove(full)


def _parse_only(argv) -> set:
    """Return the set of library folders named by ``--only a,b`` (or empty)."""
    folders: set[str] = set()
    for i, arg in enumerate(argv):
        value = ""
        if arg.startswith("--only="):
            value = arg.split("=", 1)[1]
        elif arg == "--only" and i + 1 < len(argv):
            value = argv[i + 1]
        folders.update(f.strip().lower() for f in value.split(",") if f.strip())
    return folders


def _existing_library(folder: str):
    """Return the library description of an already-generated folder, or None.

    Used with ``--only`` so the master indexes and consolidated exports still
    cover every library, not just the ones scraped in this run.
    """
    path = os.path.join(DATA_DIR, folder, "json", "_index.json")
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh).get("library") or folder
    except (OSError, ValueError):
        return None


def main() -> int:
    argv = sys.argv[1:]
    no_enrich = "--no-enrich" in argv
    force_dl = "--refresh-docs" in argv
    only = _parse_only(argv)

    unknown = only - {folder for (folder, _i, _e, _p) in LIBRARIES}
    if unknown:
        print(f"Unknown library folder(s) for --only: {sorted(unknown)}")
        print(f"Known: {[f for (f, _i, _e, _p) in LIBRARIES]}")
        return 2

    selected = [lib for lib in LIBRARIES if not only or lib[0] in only]

    index = None
    if not no_enrich:
        prefixes = {p for (_f, _i, _e, p) in selected if p}
        print("Building MS Learn description index...")
        index = mslearn_docs.build_index(prefixes, force_download=force_dl)
        if len(index) == 0:
            print("  (no descriptions available; writing reference only)")

    total = 0
    libs_done = 0
    done_libs: list[tuple[str, str]] = []
    for folder, include, exclude, doc_prefix in LIBRARIES:
        if only and folder not in only:
            # Keep previously generated output in the master indexes.
            desc = _existing_library(folder)
            if desc:
                done_libs.append((folder, desc))
            continue
        tlb, desc = load_typelib(include, exclude)
        if tlb is None:
            print(f"[skip] No registered type library matched '{folder}' "
                  f"(include={include}).")
            continue
        lib_doc = tlb.GetDocumentation(-1)
        lib_desc = desc or lib_doc[1] or lib_doc[0] or folder
        output_dir = os.path.join(DATA_DIR, folder)
        _clean_stale_flat(output_dir)
        print(f"Scraping {lib_desc} -> reference/{folder}/ "
              f"({tlb.GetTypeInfoCount()} type entries)...")
        written = scrape_typelib(tlb, lib_desc, output_dir, doc_prefix, index)
        print(f"  Wrote {written} types to reference/{folder}/md/ and "
              f"reference/{folder}/json/")
        total += written
        libs_done += 1
        done_libs.append((folder, lib_desc))

    if libs_done == 0:
        print("No matching type libraries were found. Is Office installed?")
        return 1

    print("Building master indexes (index.json, members.json)...")
    names = write_master_indexes(DATA_DIR, done_libs)
    print(f"  Indexed {names} distinct member names.")

    print(f"Building consolidated exports "
          f"(reference/{consolidate_reference.OUTPUT_SUBDIR}/)...")
    consolidate_reference.build(DATA_DIR)

    print(f"\nDone. {total} reference files written across "
          f"{libs_done} libraries.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
