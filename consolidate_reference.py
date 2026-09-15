"""Build the consolidated, per-application exports under ``reference/consolidated/``.

The per-type files under ``reference/<library>/`` are the canonical data. This
module folds them into a small number of large files, one set per host
application, for people who want a single file to search, diff, or feed to a
model:

    <app>.json / <app>.md              every type in the application
    <app>_constants.json / .md         every enumeration and module constant
    <app>_properties.json / .md        every property of every object

``render_type_md`` is also the Markdown renderer used by the scraper for the
per-type files, so both layouts are generated from one implementation and
cannot drift apart.

Standard library only: this runs anywhere the generated JSON is present, with
no need for Windows, Office, or pywin32.

Usage:
    python consolidate_reference.py
"""

from __future__ import annotations

import json
import os
import re
import sys

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(ROOT_DIR, "reference")
OUTPUT_SUBDIR = "consolidated"

SOURCE_URL = "https://github.com/WilliamSmithEdward/pyVBAReference"

# Host applications, in the order they appear in the exports. Every library
# folder not claimed here is grouped into the trailing "shared" set, so adding
# a library to the scraper cannot silently drop it from the exports.
APPLICATIONS = [
    ("excel", "Excel", ("excel",),
     "The Excel object model: workbooks, worksheets, ranges, charts and "
     "pivot tables."),
    ("powerpoint", "PowerPoint", ("powerpoint",),
     "The PowerPoint object model: presentations, slides, shapes and "
     "animation."),
    ("word", "Word", ("word",),
     "The Word object model: documents, ranges, selections, tables and "
     "styles."),
    ("access", "Access", ("access",),
     "The Access object model: forms, reports, controls and the DoCmd "
     "action model."),
]

SHARED_KEY = "shared"
SHARED_TITLE = "Shared libraries"
SHARED_DESCRIPTION = (
    "Libraries any host application can reference: the Office UI objects, the "
    "VBA language itself, user forms, ADO, and the other common "
    "Tools > References entries."
)

# Mirrors ``_INVALID`` in vba_reference/_data.py so type names map to the same
# file names on both sides.
_INVALID = re.compile(r'[<>:"/\\|?*]')


def safe_filename(name: str) -> str:
    """Return the on-disk file stem used for a type ``name``."""
    return _INVALID.sub("_", name)


# --------------------------------------------------------------------------- #
# Markdown rendering (shared with the per-type scraper output)
# --------------------------------------------------------------------------- #

def _member_text(member: dict) -> str:
    """Rebuild the one-line display form of a member."""
    if member.get("kind") == "property":
        ptype = member.get("type") or ""
        return (member["name"] + (f" As {ptype}" if ptype else "")
                + f"  ({member.get('access', '')})")
    return member.get("signature") or member["name"]


def _emit_members(lines: list, members) -> None:
    for m in members:
        line = f"- `{_member_text(m)}`"
        if m.get("description"):
            line += f"  \n  {m['description']}"
        lines.append(line)
        for p in m.get("parameters", []):
            if p.get("description"):
                opt = "optional" if p.get("optional") else "required"
                t = f" As {p['type']}" if p.get("type") else ""
                lines.append(
                    f"    - `{p['name']}{t}` ({opt}): {p['description']}")


def _emit_remarks_example(lines: list, remarks: str, example: str) -> None:
    if remarks:
        lines.append(f"**Remarks:** {remarks}")
        lines.append("")
    if example:
        lines.append("**Example:**")
        lines.append("")
        lines.append("```vba")
        lines.append(example)
        lines.append("```")
        lines.append("")


def _emit_constants(lines: list, constants) -> None:
    """Enumeration constants: name, value, and any description."""
    for c in constants:
        line = f"- `{c['name']}`"
        if c.get("value") is not None:
            line += f" = {c['value']}"
        if c.get("description"):
            line += f"  \n  {c['description']}"
        lines.append(line)


def _emit_module_constants(lines: list, constants) -> None:
    """Module constants: name, declared type, and value (no descriptions)."""
    for c in constants:
        t = f" As {c['type']}" if c.get("type") else ""
        v = f" = {c['value']}" if c.get("value") is not None else ""
        lines.append(f"- `{c['name']}{t}{v}`")


def render_type_md(data: dict, level: int = 1) -> list:
    """Render one type's JSON entry as Markdown lines.

    ``level`` is the heading level of the type name (1 for a standalone file,
    deeper when the type is nested inside a consolidated document). Member
    sections sit one level below it.
    """
    head = "#" * level
    sub = "#" * (level + 1)
    kind = data.get("kind", "")
    lines = [f"{head} {data['name']}", ""]
    lines.append(f"**Type:** {kind}  ")
    lines.append(f"**Library:** {data.get('library', '')}  ")
    if data.get("guid"):
        lines.append(f"**GUID:** {data['guid']}  ")
    lines.append("")
    if data.get("description"):
        lines.append(data["description"])
        lines.append("")

    if kind == "Enumeration":
        constants = data.get("constants", [])
        lines.append(f"{sub} Constants ({len(constants)})")
        lines.append("")
        _emit_constants(lines, constants)
        lines.append("")
        return lines

    if kind == "Module":
        constants = data.get("constants", [])
        functions = data.get("functions", [])
        if constants:
            lines.append(f"{sub} Constants ({len(constants)})")
            lines.append("")
            _emit_module_constants(lines, constants)
            lines.append("")
        if functions:
            lines.append(f"{sub} Functions ({len(functions)})")
            lines.append("")
            _emit_members(lines, functions)
            lines.append("")
        return lines

    _emit_remarks_example(lines, data.get("remarks", ""), data.get("example", ""))
    groups = (("Properties", data.get("properties", [])),
              ("Methods", data.get("methods", [])),
              ("Events", data.get("events", [])))
    for label, members in groups:
        if members:
            lines.append(f"{sub} {label} ({len(members)})")
            lines.append("")
            _emit_members(lines, members)
            lines.append("")
    if not any(members for _label, members in groups):
        lines.append("_No public members._")
        lines.append("")
    return lines


# --------------------------------------------------------------------------- #
# Loading the generated per-type data
# --------------------------------------------------------------------------- #

def _read_json(path: str) -> dict:
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def load_library(data_dir: str, folder: str, types) -> list:
    """Load every per-type JSON file of one library, in catalog order."""
    json_dir = os.path.join(data_dir, folder, "json")
    out = []
    for entry in types:
        path = os.path.join(json_dir, safe_filename(entry["name"]) + ".json")
        try:
            out.append(_read_json(path))
        except OSError:
            print(f"  [warn] missing {path}")
    return out


def applications(index: dict) -> list:
    """Group the libraries of ``index.json`` into application sets."""
    by_folder = {lib["folder"]: lib for lib in index["libraries"]}
    claimed = set()
    groups = []
    for key, title, folders, description in APPLICATIONS:
        present = [by_folder[f] for f in folders if f in by_folder]
        if not present:
            continue
        claimed.update(lib["folder"] for lib in present)
        groups.append({"key": key, "title": title,
                       "description": description, "libraries": present})
    shared = [lib for lib in index["libraries"]
              if lib["folder"] not in claimed]
    if shared:
        groups.append({"key": SHARED_KEY, "title": SHARED_TITLE,
                       "description": SHARED_DESCRIPTION, "libraries": shared})
    return groups


# --------------------------------------------------------------------------- #
# Export builders
# --------------------------------------------------------------------------- #

def _header_md(title: str, description: str, summary: list,
               siblings: str) -> list:
    """The preamble every export shares: title, blurb, counts, provenance."""
    lines = [f"# {title}", "", description, ""]
    for item in summary:
        lines.append(f"- {item}")
    lines.append("")
    lines.append(f"Generated by pyVBAReference ({SOURCE_URL}) from the "
                 f"registered COM type libraries. {siblings}")
    lines.append("")
    return lines


def build_full(group: dict, loaded: dict) -> tuple:
    """Every type of the application: one JSON object and one Markdown doc."""
    key = group["key"]
    libraries = []
    md_body = []
    total = 0
    for lib in group["libraries"]:
        types = loaded[lib["folder"]]
        total += len(types)
        libraries.append({"folder": lib["folder"], "library": lib["library"],
                          "type_count": len(types), "types": types})
        md_body.append(f"## {lib['library']}")
        md_body.append("")
        md_body.append(f"{len(types):,} types, in alphabetical order.")
        md_body.append("")
        for data in types:
            md_body.extend(render_type_md(data, level=3))

    doc = {
        "schema": "pyvbareference/application/1",
        "application": group["title"],
        "key": key,
        "description": group["description"],
        "source": SOURCE_URL,
        "library_count": len(libraries),
        "type_count": total,
        "libraries": libraries,
    }
    md = _header_md(
        f"{group['title']} VBA object model", group["description"],
        [f"**Libraries:** {len(libraries)}", f"**Types:** {total:,}"],
        f"See `{key}_constants.md` for the enumerations alone and "
        f"`{key}_properties.md` for the properties alone.",
    ) + md_body
    return doc, md


def build_constants(group: dict, loaded: dict) -> tuple:
    """Every enumeration and module constant of the application."""
    key = group["key"]
    libraries = []
    md_body = []
    groups_total = 0
    consts_total = 0
    for lib in group["libraries"]:
        entries = []
        for data in loaded[lib["folder"]]:
            if data.get("kind") not in ("Enumeration", "Module"):
                continue
            constants = data.get("constants", [])
            if not constants:
                continue
            entries.append({"name": data["name"], "kind": data["kind"],
                            "description": data.get("description", ""),
                            "constant_count": len(constants),
                            "constants": constants})
        if not entries:
            continue
        groups_total += len(entries)
        consts_total += sum(e["constant_count"] for e in entries)
        libraries.append({"folder": lib["folder"], "library": lib["library"],
                          "enumeration_count": len(entries),
                          "enumerations": entries})
        md_body.append(f"## {lib['library']}")
        md_body.append("")
        md_body.append(f"{len(entries):,} enumerations and modules.")
        md_body.append("")
        for entry in entries:
            md_body.append(f"### {entry['name']}")
            md_body.append("")
            md_body.append(f"**Type:** {entry['kind']}  ")
            md_body.append(f"**Constants:** {entry['constant_count']}  ")
            md_body.append("")
            if entry["description"]:
                md_body.append(entry["description"])
                md_body.append("")
            if entry["kind"] == "Module":
                _emit_module_constants(md_body, entry["constants"])
            else:
                _emit_constants(md_body, entry["constants"])
            md_body.append("")

    doc = {
        "schema": "pyvbareference/constants/1",
        "application": group["title"],
        "key": key,
        "source": SOURCE_URL,
        "enumeration_count": groups_total,
        "constant_count": consts_total,
        "libraries": libraries,
    }
    md = _header_md(
        f"{group['title']} VBA constants", group["description"],
        [f"**Enumerations and modules:** {groups_total:,}",
         f"**Constants:** {consts_total:,}"],
        f"Constant values only; see `{key}.md` for the full object model.",
    ) + md_body
    return doc, md


def build_properties(group: dict, loaded: dict) -> tuple:
    """Every property of every object of the application."""
    key = group["key"]
    libraries = []
    md_body = []
    types_total = 0
    props_total = 0
    for lib in group["libraries"]:
        entries = []
        for data in loaded[lib["folder"]]:
            properties = data.get("properties", [])
            if not properties:
                continue
            entries.append({"name": data["name"], "kind": data.get("kind", ""),
                            "property_count": len(properties),
                            "properties": properties})
        if not entries:
            continue
        types_total += len(entries)
        props_total += sum(e["property_count"] for e in entries)
        libraries.append({"folder": lib["folder"], "library": lib["library"],
                          "type_count": len(entries), "types": entries})
        md_body.append(f"## {lib['library']}")
        md_body.append("")
        md_body.append(f"{len(entries):,} objects with properties.")
        md_body.append("")
        for entry in entries:
            md_body.append(f"### {entry['name']}")
            md_body.append("")
            md_body.append(f"**Type:** {entry['kind']}  ")
            md_body.append(f"**Properties:** {entry['property_count']}  ")
            md_body.append("")
            _emit_members(md_body, entry["properties"])
            md_body.append("")

    doc = {
        "schema": "pyvbareference/properties/1",
        "application": group["title"],
        "key": key,
        "source": SOURCE_URL,
        "type_count": types_total,
        "property_count": props_total,
        "libraries": libraries,
    }
    md = _header_md(
        f"{group['title']} VBA properties", group["description"],
        [f"**Objects:** {types_total:,}", f"**Properties:** {props_total:,}"],
        f"Properties only; see `{key}.md` for methods, events and "
        f"parameter documentation.",
    ) + md_body
    return doc, md


BUILDERS = (("", build_full), ("_constants", build_constants),
            ("_properties", build_properties))


# --------------------------------------------------------------------------- #
# Writing
# --------------------------------------------------------------------------- #

def _write_json(path: str, doc: dict) -> int:
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(doc, fh, indent=2)
        fh.write("\n")
    return os.path.getsize(path)


def _write_md(path: str, lines) -> int:
    text = "\n".join(lines)
    if not text.endswith("\n"):
        text += "\n"
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(text)
    return os.path.getsize(path)


def _size(num: int) -> str:
    return f"{num / 1048576:.1f} MB" if num >= 1048576 else f"{num // 1024} KB"


def _write_index(out_dir: str, rows) -> None:
    def cell(key: str, suffix: str) -> str:
        return (f"[{key}{suffix}.md]({key}{suffix}.md) / "
                f"[json]({key}{suffix}.json)")

    lines = [
        "# Consolidated exports", "",
        "One set of files per host application, folded together from the "
        "per-type files in `reference/<library>/`. Same data, fewer files.", "",
        "| Application | Object model | Constants | Properties |",
        "| ----------- | ------------ | --------- | ---------- |",
    ]
    for row in rows:
        key = row["key"]
        lines.append(f"| {row['title']} | {cell(key, '')} | "
                     f"{cell(key, '_constants')} | "
                     f"{cell(key, '_properties')} |")
    lines.append("")
    lines.append("Counts:")
    lines.append("")
    for row in rows:
        lines.append(f"- **{row['title']}** - {row['type_count']:,} types, "
                     f"{row['constant_count']:,} constants, "
                     f"{row['property_count']:,} properties")
    lines.append("")
    with open(os.path.join(out_dir, "_index.md"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines))


def build(data_dir: str = DATA_DIR, quiet: bool = False) -> int:
    """Write every consolidated export. Returns the number of files written."""
    index = _read_json(os.path.join(data_dir, "index.json"))
    out_dir = os.path.join(data_dir, OUTPUT_SUBDIR)
    os.makedirs(out_dir, exist_ok=True)

    written = 0
    rows = []
    for group in applications(index):
        loaded = {lib["folder"]: load_library(data_dir, lib["folder"],
                                              lib["types"])
                  for lib in group["libraries"]}
        docs = {}
        for suffix, builder in BUILDERS:
            doc, md = builder(group, loaded)
            docs[suffix] = doc
            stem = os.path.join(out_dir, group["key"] + suffix)
            json_bytes = _write_json(stem + ".json", doc)
            md_bytes = _write_md(stem + ".md", md)
            written += 2
            if not quiet:
                print(f"  {group['key']}{suffix}.json ({_size(json_bytes)}), "
                      f"{group['key']}{suffix}.md ({_size(md_bytes)})")
        rows.append({
            "key": group["key"], "title": group["title"],
            "type_count": docs[""]["type_count"],
            "constant_count": docs["_constants"]["constant_count"],
            "property_count": docs["_properties"]["property_count"],
        })

    _write_index(out_dir, rows)
    return written + 1


def main() -> int:
    if not os.path.isfile(os.path.join(DATA_DIR, "index.json")):
        print("No reference/index.json found. Run "
              "scrape_excel_object_model.py first.")
        return 1
    print(f"Writing consolidated exports to reference/{OUTPUT_SUBDIR}/ ...")
    written = build()
    print(f"  {written} files written.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
