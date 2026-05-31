# Outline

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208AB-0000-0000-C000-000000000046}  

Represents an outline on a worksheet.

**Example:**

```vba
Worksheets("sheet4").Outline.ShowLevels 1
```

## Properties (6)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `AutomaticStyles As Boolean  (read/write)`  
  True if the outline uses automatic styles. Read/write Boolean.
- `SummaryColumn As XlSummaryColumn  (read/write)`  
  Returns or sets the location of the summary columns in the outline. Read/write XlSummaryColumn.
- `SummaryRow As XlSummaryRow  (read/write)`  
  Returns or sets the location of the summary rows in the outline. Read/write XlSummaryRow.

## Methods (1)

- `ShowLevels([RowLevels As Variant], [ColumnLevels As Variant]) As Variant`  
  Displays the specified number of row and/or column levels of an outline.
    - `RowLevels As Variant` (optional): Specifies the number of row levels of an outline to display. If the outline has fewer levels than the number specified, Microsoft Excel displays all the levels. If this argument is 0 (zero) or is omitted, no action is taken on rows.
    - `ColumnLevels As Variant` (optional): Specifies the number of column levels of an outline to display. If the outline has fewer levels than the number specified, Excel displays all the levels. If this argument is 0 (zero) or is omitted, no action is taken on columns.
