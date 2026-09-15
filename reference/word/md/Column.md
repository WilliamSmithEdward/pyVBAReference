# Column

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002094F-0000-0000-C000-000000000046}  

Represents a single table column. The Column object is a member of the Columns collection. The Columns collection includes all the columns in a table, selection, or range.

**Remarks:** Use Columns (Index), where Index is the index number, to return a single Column object. The index number represents the position of the column in the Columns collection (counting from left to right). The following example selects column one in table one in the active document. Use the Column property with a Cell object to return a Column object. The following example deletes the text in cell one, inserts new text, and then sorts the entire column. Use the Add method to add a column to a table. The following example adds a column to the first table in the active document, and then it makes the column widths equal. Remarks Use the Information property with a Selection object to return the current column number. The following example selects the current column and then displays the column number in a message box.

## Properties (15)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Column object.
- `Width As Single  (read/write)`  
  Returns or sets the width of the specified column, in points. Read/write Long.
- `IsFirst As Boolean  (read-only)`  
  True if the specified column or row is the first one in the table. Read-only Boolean.
- `IsLast As Boolean  (read-only)`  
  True if the specified column or row is the last one in the table. Read-only Boolean.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
- `Cells As Cells  (read-only)`  
  Returns a Cells collection that represents the table cells in a table column. Read-only.
- `Borders As Borders  (read/write)`  
  Returns a Borders collection that represents all the borders for the specified object.
- `Shading As Shading  (read-only)`  
  Returns a Shading object that refers to the shading formatting for the specified column.
- `Next As Column  (read-only)`  
  Returns the next column in a collection of table columns. Read-only.
- `Previous As Column  (read-only)`  
  Returns the previous column in a collection of table columns. Read-only.
- `NestingLevel As Long  (read-only)`  
  Returns the nesting level of the specified column. Read-only Long.
- `PreferredWidth As Single  (read/write)`  
  Returns or sets the preferred width (in points or as a percentage of the window width) for the specified column. Read/write Single.
- `PreferredWidthType As WdPreferredWidthType  (read/write)`  
  Returns or sets the preferred unit of measurement to use for the width of the specified table column. Read/write WdPreferredWidthType.

## Methods (5)

- `Select()`  
  Selects the specified table column.
- `Delete()`  
  Deletes the specified column.
- `SetWidth(ColumnWidth As Single, RulerStyle As WdRulerStyle)`  
  Sets the width of a column in a table.
    - `ColumnWidth As Single` (required): The width of the specified column or columns, in points.
    - `RulerStyle As WdRulerStyle` (required): Controls the way Word adjusts cell widths.
- `AutoFit()`  
  Changes the width of a table column to accommodate the width of the text without changing the way text wraps in the cells.
- `Sort([ExcludeHeader As Variant], [SortFieldType As Variant], [SortOrder As Variant], [CaseSensitive As Variant], [BidiSort As Variant], [IgnoreThe As Variant], [IgnoreKashida As Variant], [IgnoreDiacritics As Variant], [IgnoreHe As Variant], [LanguageID As Variant])`  
  Sorts the specified table column.
    - `ExcludeHeader As Variant` (optional): True to exclude the first row or paragraph header from the sort operation. The default value is False.
    - `SortFieldType As Variant` (optional): The sort type for the column. Can be one of the WdSortFieldType constants.
    - `SortOrder As Variant` (optional): The sorting order to use for the column. Can be one WdSortOrder constant.
    - `CaseSensitive As Variant` (optional): True to sort with case sensitivity. The default value is False.
    - `BidiSort As Variant` (optional): True to sort based on right-to-left language rules. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `IgnoreThe As Variant` (optional): True to ignore the Arabic character alef lam when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `IgnoreKashida As Variant` (optional): True to ignore kashidas when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `IgnoreDiacritics As Variant` (optional): True to ignore bidirectional control characters when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `IgnoreHe As Variant` (optional): True to ignore the Hebrew character he when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `LanguageID As Variant` (optional): Specifies the sorting language. Can be one of the WdLanguageID constants.
