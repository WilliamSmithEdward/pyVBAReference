# Agentic LLM Primer

A guide for AI agents on how to use this VBA reference library to write correct,
well-grounded VBA (Visual Basic for Applications) code.

## What this repository is

This repository is a **machine-generated reference of the VBA object models** for
ten commonly referenced COM type libraries. Every public type (class, interface,
enumeration, and module) is exported in two parallel forms:

- `<library>/md/` - human-readable Markdown, one file per type, plus `_index.md`.
- `<library>/json/` - machine-readable JSON, one file per type, plus `_index.json`.

The two folders contain the same information. **Prefer the JSON when reasoning
programmatically** (stable schema, easy to parse); use the Markdown when you only
need to read or quote documentation.

Member signatures, return types, parameter lists, property access modes, enum
values, remarks, and examples are introspected directly from the registered COM
type libraries and then enriched with descriptions scraped from the official
Microsoft Learn VBA documentation.

## Libraries available

| Folder      | Library                                              | Types |
| ----------- | ---------------------------------------------------- | ----- |
| `excel`     | Microsoft Excel 16.0 Object Library                  | 1028  |
| `office`    | Microsoft Office 16.0 Object Library                 | 510   |
| `msforms`   | Microsoft Forms 2.0 Object Library                   | 166   |
| `adodb`     | Microsoft ActiveX Data Objects 6.1 Library           | 110   |
| `msxml`     | Microsoft XML, v6.0                                  | 101   |
| `vbide`     | Microsoft Visual Basic for Applications Extensibility | 46    |
| `scripting` | Microsoft Scripting Runtime                          | 28    |
| `vba`       | Visual Basic For Applications                        | 26    |
| `stdole`    | OLE Automation                                       | 11    |
| `winhttp`   | Microsoft WinHTTP Services, version 5.1              | 7     |

The `vba` library holds the **language built-ins** - global functions like
`MsgBox`, `Format`, `CStr`, `Left`, and intrinsic constants like `vbCrLf` - grouped
into modules (`Interaction`, `Strings`, `Conversion`, `DateTime`, `Math`,
`FileSystem`, `Information`, `Financial`, `Constants`, ...). These are not methods
on an object; they are callable from anywhere in VBA.

## How to find something

1. **Don't know which type defines a member?** Read the repo-root `members.json`
   (see below). It maps every member name to the types that define it across all
   libraries - the fastest way to answer "where is `SaveAs`?".
2. **Want the catalog of everything?** Read the repo-root `index.json` (see below)
   for all libraries and their types in a single file.
3. **Know the type name?** Open `<library>/json/<TypeName>.json` (or `.md`). File
   names match the type name (with characters unsafe for file systems replaced).
4. **Browsing one library?** Read `<library>/json/_index.json` for the full list of
   `{ "name", "kind" }` entries, or `<library>/md/_index.md` for a linked index.
5. **Looking for a global function** (e.g. `MsgBox`)? It is a function inside a
   module in the `vba` library - e.g. `MsgBox` lives in `vba/json/Interaction.json`
   under the `functions` array. `members.json` will point you straight to it.
6. **Looking for a constant value** (e.g. what number is `xlCSV`)? Open the relevant
   enumeration file, e.g. `excel/json/XlFileFormat.json`, and read its `constants`.

## JSON schema

Every JSON file has at least `name`, `kind`, and `library`. The remaining fields
depend on `kind`.

### Class / interface (`kind`: `"Class"`, `"Dispatch Interface"`, `"Interface"`)

```json
{
  "name": "Worksheet",
  "kind": "Class",
  "guid": "{00020820-0000-0000-C000-000000000046}",
  "library": "Microsoft Excel 16.0 Object Library",
  "description": "Represents a worksheet.",
  "remarks": "The Worksheet object is a member of the Worksheets collection ...",
  "example": "Worksheets(1).Visible = False",
  "properties": [ /* property objects */ ],
  "methods":    [ /* method objects   */ ],
  "events":     [ /* method objects   */ ]
}
```

**Property object:**

```json
{
  "name": "Application",
  "kind": "property",
  "type": "Application",
  "access": "read-only",
  "description": "When used without an object qualifier, this property returns ..."
}
```

`access` is one of `"read-only"`, `"write-only"`, or `"read/write"`.

**Method / event / function object:**

```json
{
  "name": "PrintOut",
  "kind": "method",
  "signature": "PrintOut([From As Variant], [To As Variant], [Copies As Variant], ...)",
  "description": "Prints the object.",
  "returns": "void",
  "parameters": [
    {
      "name": "From",
      "type": "Variant",
      "optional": true,
      "description": "The number of the page at which to start printing ..."
    }
  ]
}
```

- `returns` is the VBA return type, or `"void"` for a `Sub` (no return value).
- `signature` is the exact VBA call signature. Square brackets mark optional
  parameters. The same optionality is also encoded per parameter via `optional`.

### Enumeration (`kind`: `"Enumeration"`)

```json
{
  "name": "XlFileFormat",
  "kind": "Enumeration",
  "library": "Microsoft Excel 16.0 Object Library",
  "description": "Specifies the file format when saving the worksheet.",
  "constants": [
    { "name": "xlCSV", "value": 6, "description": "CSV" }
  ]
}
```

### Module (`kind`: `"Module"`)

Holds global functions and intrinsic constants (mostly in the `vba` library).

```json
{
  "name": "Interaction",
  "kind": "Module",
  "library": "Visual Basic For Applications",
  "description": "...",
  "functions": [ /* method/function objects, same shape as above */ ],
  "constants": [ { "name": "...", "value": 0, "type": "..." } ]
}
```

### Per-library index (`<library>/json/_index.json`)

```json
{
  "library": "Microsoft Excel 16.0 Object Library",
  "types": [ { "name": "Worksheet", "kind": "Class" } ]
}
```

## Repo-root master indexes

Two files at the repository root span all libraries and are regenerated with the
rest of the data.

### `index.json` - master catalog

```json
{
  "total_libraries": 10,
  "total_types": 2033,
  "libraries": [
    {
      "folder": "excel",
      "library": "Microsoft Excel 16.0 Object Library",
      "type_count": 1028,
      "types": [ { "name": "Worksheet", "kind": "Class" } ]
    }
  ]
}
```

### `members.json` - cross-library member lookup

Maps every member name (property / method / event / function / constant) to the
types that define it. Use it to resolve a member to its owning type and library
in one read.

```json
{
  "total_names": 9948,
  "members": {
    "SaveAs": [
      { "library": "excel", "type": "Chart",     "kind": "method" },
      { "library": "excel", "type": "Workbook",  "kind": "method" }
    ]
  }
}
```

Keys are the exact member names; resolve `library` + `type` to the per-type file
at `<library>/json/<type>.json` for the full signature and parameter docs.

## Rules of engagement for agents

- **Ground every API call.** Before emitting VBA that calls a method, property, or
  function, verify the exact name, parameter order, optionality, and return type
  against the relevant JSON file. Do not rely on memory for signatures.
- **Respect access modes.** Never assign to a property whose `access` is
  `"read-only"`. Only read a `"write-only"` property's effect through other means.
- **Match parameter order and types.** Use the `parameters` array. Optional
  parameters (`"optional": true`) may be omitted or passed by name
  (`Range:=..., Style:=...`).
- **Use real constant values.** When a parameter expects an enum, reference the
  named constant (e.g. `xlCSV`) rather than a magic number; the numeric `value` is
  available if you must inline it.
- **Distinguish Sub from Function.** If `returns` is `"void"`, call it as a
  statement (`obj.PrintOut`), not in an expression. Otherwise it returns a value.
- **Global vs. member.** Language built-ins (`MsgBox`, `Format`, ...) come from the
  `vba` library modules and are called bare. Everything else is a member reached
  through an object (`Worksheet.Protect`, `Workbook.SaveAs`, ...).
- **Quote documentation faithfully.** The `description`, `remarks`, and `example`
  fields are sourced from Microsoft Learn; treat them as authoritative and quote
  rather than paraphrase when accuracy matters.

## Quick lookup recipes

Find a member's signature in Python:

```python
import json
wb = json.load(open("excel/json/Worksheet.json", encoding="utf-8"))
protect = next(m for m in wb["methods"] if m["name"] == "Protect")
print(protect["signature"])
for p in protect["parameters"]:
    print(p["name"], p["type"], "optional" if p["optional"] else "required")
```

Resolve an enum constant's value:

```python
import json
fmt = json.load(open("excel/json/XlFileFormat.json", encoding="utf-8"))
print({c["name"]: c["value"] for c in fmt["constants"]}["xlCSV"])  # -> 6
```

Locate a global function across the `vba` library:

```python
import json
mem = json.load(open("members.json", encoding="utf-8"))
for hit in mem["members"]["MsgBox"]:
    print(hit)  # -> {'library': 'vba', 'type': 'Interaction', 'kind': 'function'}
```

## Regenerating the reference

The data is produced by `scrape_excel_object_model.py`, which introspects the
COM type libraries registered on a Windows machine with Microsoft Office installed
and enriches them from the Microsoft Learn VBA-Docs corpus. To rebuild:

```powershell
.venv\Scripts\python.exe scrape_excel_object_model.py
```

Useful flags: `--no-enrich` (skip Microsoft Learn descriptions; signatures only) and
`--refresh-docs` (force re-download of the documentation corpus). Regenerating
overwrites the `md/` and `json/` folders and removes any stale flat files from
earlier layouts.
