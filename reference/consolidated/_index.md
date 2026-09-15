# Consolidated exports

One set of files per host application, folded together from the per-type files in `reference/<library>/`. Same data, fewer files.

| Application | Object model | Constants | Properties |
| ----------- | ------------ | --------- | ---------- |
| Excel | [excel.md](excel.md) / [json](excel.json) | [excel_constants.md](excel_constants.md) / [json](excel_constants.json) / [csv](excel_constants.csv) | [excel_properties.md](excel_properties.md) / [json](excel_properties.json) / [csv](excel_properties.csv) |
| PowerPoint | [powerpoint.md](powerpoint.md) / [json](powerpoint.json) | [powerpoint_constants.md](powerpoint_constants.md) / [json](powerpoint_constants.json) / [csv](powerpoint_constants.csv) | [powerpoint_properties.md](powerpoint_properties.md) / [json](powerpoint_properties.json) / [csv](powerpoint_properties.csv) |
| Word | [word.md](word.md) / [json](word.json) | [word_constants.md](word_constants.md) / [json](word_constants.json) / [csv](word_constants.csv) | [word_properties.md](word_properties.md) / [json](word_properties.json) / [csv](word_properties.csv) |
| Access | [access.md](access.md) / [json](access.json) | [access_constants.md](access_constants.md) / [json](access_constants.json) / [csv](access_constants.csv) | [access_properties.md](access_properties.md) / [json](access_properties.json) / [csv](access_properties.csv) |
| Shared libraries | [shared.md](shared.md) / [json](shared.json) | [shared_constants.md](shared_constants.md) / [json](shared_constants.json) / [csv](shared_constants.csv) | [shared_properties.md](shared_properties.md) / [json](shared_properties.json) / [csv](shared_properties.csv) |

Counts:

- **Excel** - 1,028 types, 2,328 constants, 9,896 properties
- **PowerPoint** - 339 types, 1,480 constants, 2,009 properties
- **Word** - 750 types, 3,756 constants, 4,427 properties
- **Access** - 289 types, 1,595 constants, 5,010 properties
- **Shared libraries** - 1,004 types, 3,836 constants, 5,201 properties

## The CSV tables

The same properties and constants as flat tables, for importing into a spreadsheet.

`<app>_properties.csv`:

| Column | Meaning |
| ------ | ------- |
| `Object` | the type that owns the property |
| `Property` | the property name, as written in code |
| `Property split` | the same name split at capitals, `ActiveCell` -> `Active Cell` |
| `Type` | what the property returns |
| `Access` | `R_O` read-only, `V` settable, `W_O` write-only |

`<app>_constants.csv`:

| Column | Meaning |
| ------ | ------- |
| `Owner` | the enumeration or module that declares it |
| `Kind` | `Enumeration` or `Module` |
| `Constant` | the constant name, as written in code |
| `Value` | the value as a VB6 literal: `6`, `-4104`, `"PDF Format (*.pdf)"`, `Chr(13) & Chr(10)` |
| `Description` | from Microsoft Learn, where there is one |
