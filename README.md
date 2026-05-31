# vba_reference_library

A complete, machine-generated reference of the VBA (Visual Basic for Applications)
object models for ten common COM type libraries. Every public type is exported as
both human-readable Markdown and machine-readable JSON.

## Layout

```
<library>/
  md/     one .md per type   + _index.md
  json/   one .json per type + _index.json
```

Both folders hold the same data: signatures, return types, parameter lists,
property access modes, enum values, remarks, and examples - introspected from the
registered COM type libraries and enriched with descriptions from Microsoft Learn.

## Libraries

| Folder      | Library                                              | Types |
| ----------- | ---------------------------------------------------- | ----- |
| `excel`     | Microsoft Excel 16.0 Object Library                  | 1028  |
| `office`    | Microsoft Office 16.0 Object Library                 | 510   |
| `msforms`   | Microsoft Forms 2.0 Object Library                   | 166   |
| `adodb`     | Microsoft ActiveX Data Objects 6.1 Library           | 110   |
| `msxml`     | Microsoft XML, v6.0                                  | 101   |
| `vbide`     | Microsoft Visual Basic for Applications Extensibility | 46    |
| `scripting` | Microsoft Scripting Runtime                          | 28    |
| `vba`       | Visual Basic For Applications (language built-ins)   | 26    |
| `stdole`    | OLE Automation                                       | 11    |
| `winhttp`   | Microsoft WinHTTP Services, version 5.1              | 7     |

VBA language built-ins (`MsgBox`, `Format`, `CStr`, `vbCrLf`, ...) live in the
`vba` library, grouped into modules such as `Interaction` and `Strings`.

## Finding something

- Don't know which type owns a member? Check `members.json` (repo root) - it maps
  every member name to the types that define it.
- Want the full catalog? See `index.json` (repo root) - all libraries and types.
- Know the type name? Open `<library>/md/<TypeName>.md`.
- Browsing one library? Start at `<library>/md/_index.md`.
- A global function (e.g. `MsgBox`)? It's in a module under `vba/` -
  `MsgBox` is in `vba/md/Interaction.md`.
- A constant's value (e.g. `xlCSV`)? See the enum file, e.g.
  `excel/md/XlFileFormat.md`.

## Regenerating

Requires Windows with Microsoft Office installed and `pywin32`:

```powershell
.venv\Scripts\python.exe scrape_excel_object_model.py
```

Flags: `--no-enrich` (signatures only, skip Microsoft Learn text) and
`--refresh-docs` (force re-download of the documentation corpus).

## For AI agents

See [agentic_llm_primer.md](agentic_llm_primer.md) for the JSON schema and
guidance on grounding VBA code against this reference.
