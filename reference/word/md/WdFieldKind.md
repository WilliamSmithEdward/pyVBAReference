# WdFieldKind

**Type:** Enumeration  
**Library:** Microsoft Word 16.0 Object Library  

Specifies the type of field for a Field object.

## Constants (4)

- `wdFieldKindNone` = 0  
  An invalid field (for example, a pair of field characters with nothing inside).
- `wdFieldKindHot` = 1  
  A field that's automatically updated each time it is displayed or each time the page is reformatted, but which can also be manually updated (for example, INCLUDEPICTURE or FORMDROPDOWN).
- `wdFieldKindWarm` = 2  
  A field that can be updated and has a result. This type includes fields that are automatically updated when the source changes and fields that can be manually updated (for example, DATE or INCLUDETEXT).
- `wdFieldKindCold` = 3  
  A field that does not have a result, for example, an Index Entry (XE), Table of Contents Entry (TC), or Private field.
