# pyVBAReference

A complete, machine-generated reference of the VBA (Visual Basic for Applications)
object models for Excel, PowerPoint, Word and Access, plus the nine shared COM
type libraries every host can reference. Every public type is exported as both
human-readable Markdown and machine-readable JSON.

## Layout

All generated data lives under `reference/` to keep the repo root clean:

```
reference/
  agentic_llm_primer.md  guide for AI agents (schema + grounding)
  index.json             master catalog of every library and type
  members.json           member name -> the types that define it
  <library>/
    md/     one .md per type   + _index.md
    json/   one .json per type + _index.json
  consolidated/          the same data folded into one set of files per app
    <app>.json / .md             every type
    <app>_constants.json / .md   every enumeration and module constant
    <app>_properties.json / .md  every property of every object
```

Both `md/` and `json/` hold the same data: signatures, return types, parameter
lists, property access modes, enum values, remarks, and examples - introspected
from the registered COM type libraries and enriched with descriptions from
Microsoft Learn.

## Libraries

Host applications:

| Folder       | Library                                  | Types |
| ------------ | ---------------------------------------- | ----- |
| `excel`      | Microsoft Excel 16.0 Object Library      | 1028  |
| `word`       | Microsoft Word 16.0 Object Library       | 750   |
| `powerpoint` | Microsoft PowerPoint 16.0 Object Library | 339   |
| `access`     | Microsoft Access 16.0 Object Library     | 289   |

Shared libraries, referenceable from any host:

| Folder      | Library                                              | Types |
| ----------- | ---------------------------------------------------- | ----- |
| `office`    | Microsoft Office 16.0 Object Library                 | 510   |
| `msforms`   | Microsoft Forms 2.0 Object Library                   | 166   |
| `adodb`     | Microsoft ActiveX Data Objects 6.1 Library           | 110   |
| `msxml`     | Microsoft XML, v6.0                                  | 101   |
| `vbide`     | Microsoft Visual Basic for Applications Extensibility | 45    |
| `scripting` | Microsoft Scripting Runtime                          | 28    |
| `vba`       | Visual Basic For Applications (language built-ins)   | 26    |
| `stdole`    | OLE Automation                                       | 11    |
| `winhttp`   | Microsoft WinHTTP Services, version 5.1              | 7     |

VBA language built-ins (`MsgBox`, `Format`, `CStr`, `vbCrLf`, ...) live in the
`vba` library, grouped into modules such as `Interaction` and `Strings`.

A type name that exists in several hosts (`Application`, `Range`, `Font`, ...)
resolves to the Excel one first; pass a library to pick another
(`vba.get_type("Application", "word")`).

## Finding something

- Don't know which type owns a member? Check `reference/members.json` - it maps
  every member name to the types that define it.
- Want the full catalog? See `reference/index.json` - all libraries and types.
- Know the type name? Open `reference/<library>/md/<TypeName>.md`.
- Browsing one library? Start at `reference/<library>/md/_index.md`.
- A global function (e.g. `MsgBox`)? It's in a module under `reference/vba/` -
  `MsgBox` is in `reference/vba/md/Interaction.md`.
- A constant's value (e.g. `xlCSV`)? See the enum file, e.g.
  `reference/excel/md/XlFileFormat.md`.
- Want one file instead of a folder tree? See `reference/consolidated/` below.

## Consolidated exports

`reference/consolidated/` holds the same data folded into one set of files per
application - Excel, PowerPoint, Word, Access, and `shared` for the libraries all
four have in common. Each set is three pairs of files:

| File                      | Contents                                          |
| ------------------------- | ------------------------------------------------- |
| `<app>.json` / `.md`      | every type, with methods, events and parameter docs |
| `<app>_constants.json` / `.md`  | every enumeration and module constant, with values |
| `<app>_properties.json` / `.md` | every property of every object, with type and access |

So the whole Word object model is one 3.6 MB JSON file or one 1.8 MB Markdown
file; every Excel constant is a 175 KB Markdown file. Nothing here is new data -
it is the per-type files concatenated, for grep, download, or feeding to a model.
Start at [`reference/consolidated/_index.md`](reference/consolidated/_index.md).

## Regenerating

Requires Windows with the relevant Office applications installed and `pywin32`:

```powershell
.venv\Scripts\python.exe scrape_excel_object_model.py
```

Flags: `--no-enrich` (signatures only, skip Microsoft Learn text),
`--refresh-docs` (force re-download of the documentation corpus), and
`--only word,powerpoint` (rebuild just those libraries; the master indexes and
consolidated exports still cover everything else).

The consolidated exports are rebuilt at the end of every run. To rebuild only
them - no Office or `pywin32` needed, just the generated JSON:

```powershell
python consolidate_reference.py
```

## For AI agents

See [agentic_llm_primer.md](reference/agentic_llm_primer.md) for the JSON schema and
guidance on grounding VBA code against this reference.

## Python library

The same data is exposed as an installable, typed Python package, `vba_reference`.
The JSON is bundled into the wheel, so an installed copy is self-contained; in this
repo it reads the generated folders directly.

```powershell
pip install -e .          # from this repo (editable)
# or: pip install vba-reference
```

```python
import vba_reference as vba

vba.library_names()                          # ['excel', 'powerpoint', 'word', ...]
ws = vba.get_type("Worksheet")               # TypeDoc (case-insensitive)
print(ws.remarks)
protect = ws.member("Protect")               # Member
[(p.name, p.optional) for p in protect.parameters]

vba.get_type("Document", "word")              # disambiguate by library
vba.locate_type("Application")                # every host that defines it
vba.find_members("MsgBox")                    # -> [MemberRef(library='vba', type='Interaction', ...)]
vba.find_members("SaveAs")                    # every type that defines SaveAs
vba.get_constant("XlFileFormat", "xlCSV").value   # 6
vba.get_constant("WdSaveFormat", "wdFormatPDF").value  # 17
```

Command-line interface (`vba-ref` once installed, or `python -m vba_reference`):

```powershell
vba-ref libs                     # list libraries and type counts
vba-ref where MsgBox             # where a member/type is defined
vba-ref type Worksheet           # full type entry
vba-ref member Worksheet Protect # one member with parameter docs
```

