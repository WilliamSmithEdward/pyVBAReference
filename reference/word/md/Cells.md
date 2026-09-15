# Cells

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002094A-0000-0000-C000-000000000046}  

A collection of Cell objects in a table column, table row, selection, or range.

**Remarks:** Use the Cells property to return the Cells collection. The following example formats the cells in the first row in table one in the active document to be 30 points wide. The following example returns the number of cells in the current row. Use the Add method to add a Cell object to the Cells collection. You can also use the InsertCells method of the Selection object to insert new cells. The following example adds a cell before the first cell in myTable. Use Cell (row, column), where row is the row number and column is the column number, or Cells (index), where index is the index number, to return a Cell object. The following example applies shading to the second cell in the first row in table one. The following example applies shading to the first cell in the first row. Remarks Use the Add method with the Rows or Columns collection to add a row or column of cells. The following example adds a column to the first table in the active document and then inserts numbers into the first column.

## Properties (14)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns the number of items in the Cells collection. Read-only Long.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Cells object.
- `Width As Single  (read/write)`  
  Returns or sets the width of the table cells, in points. Read/write Single.
- `Height As Single  (read/write)`  
  Returns or sets the height of the specified table cells. Read/write Single.
- `HeightRule As WdRowHeightRule  (read/write)`  
  Returns or sets a WdRowHeightRule constant that represents the rule for determining the height of the specified cells. Read/write .
- `VerticalAlignment As WdCellVerticalAlignment  (read/write)`  
  Returns or sets the vertical alignment of text in one or more cells of a table. Read/write WdCellVerticalAlignment.
- `Borders As Borders  (read/write)`  
  Returns a Borders collection that represents all the borders for the specified object.
- `Shading As Shading  (read-only)`  
  Returns a Shading object that refers to the shading formatting for the specified object.
- `NestingLevel As Long  (read-only)`  
  Returns the nesting level of the specified cells. Read-only Long.
- `PreferredWidth As Single  (read/write)`  
  Returns or sets the preferred width (in points or as a percentage of the window width) for the specified cells. Read/write Single.
- `PreferredWidthType As WdPreferredWidthType  (read/write)`  
  Returns or sets the preferred unit of measurement to use for the width of the specified cells. Read-only WdPreferredWidthType.

## Methods (10)

- `Item(Index As Long) As Cell`  
  Returns an individual Cell object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
- `Add([BeforeCell As Variant]) As Cell`  
  Returns a Cell object that represents a cell added to a table.
    - `BeforeCell As Variant` (optional): A Cell object that represents the cell that will appear immediately to the right of the new cell or cells.
- `Delete([ShiftCells As Variant])`  
  Deletes a table cell or cells and optionally controls how the remaining cells are shifted.
    - `ShiftCells As Variant` (optional): The direction in which the remaining cells are to be shifted. Can be any WdDeleteCells constant. If omitted, cells to the right of the last deleted cell are shifted left.
- `SetWidth(ColumnWidth As Single, RulerStyle As WdRulerStyle)`  
  Sets the width of columns or cells in a table.
    - `ColumnWidth As Single` (required): The width of the specified column or columns, in points.
    - `RulerStyle As WdRulerStyle` (required): Controls the way Word adjusts cell widths.
- `SetHeight(RowHeight As Variant, HeightRule As WdRowHeightRule)`  
  Sets the height of table cells.
    - `RowHeight As Variant` (required): The height of the row or rows, in points.
    - `HeightRule As WdRowHeightRule` (required): The rule for determining the height of the specified cells.
- `Merge()`  
  Merges the specified table cells with one another. The result is a single table cell.
- `Split([NumRows As Variant], [NumColumns As Variant], [MergeBeforeSplit As Variant])`  
  Splits a range of table cells.
    - `NumRows As Variant` (optional): The number of rows that the cell or group of cells is to be split into.
    - `NumColumns As Variant` (optional): The number of columns that the cell or group of cells is to be split into.
    - `MergeBeforeSplit As Variant` (optional): True to merge the cells with one another before splitting them.
- `DistributeHeight()`  
  Adjusts the height of the specified cells so that they are equal.
- `DistributeWidth()`  
  Adjusts the width of the specified cells so that they are equal.
- `AutoFit()`  
  Changes the width of a table column to accommodate the width of the text without changing the way text wraps in the cells.
