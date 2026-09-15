# Cell

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002094E-0000-0000-C000-000000000046}  

Represents a single table cell. The Cell object is a member of the Cells collection. The Cells collection represents all the cells in the specified object.

**Remarks:** Use Cell (row, column), where row is the row number and column is the column number, or Cells (index), where index is the index number, to return a Cell object. The following example applies shading to the second cell in the first row. The following example applies shading to the first cell in the first row. Use the Add method to add a Cell object to the Cells collection. You can also use the InsertCells method of the Selection object to insert new cells. The following example adds a cell before the first cell in myTable. The following example sets a range ( myRange ) that references the first two cells in the first table. After the range is set, the cells are combined by the Merge method. Remarks Use the Add method with the Rows or Columns collection to add a row or column of cells. Use the Information property with a Selection object to return the current row and column number. The following example changes the width of the first cell in the selection and then displays the cell's row number and column number.

## Properties (27)

- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that's contained in the specified object.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Cell object.
- `RowIndex As Long  (read-only)`  
  Returns the number of the row that contains the specified cell. Read-only Long.
- `ColumnIndex As Long  (read-only)`  
  Returns the number of the table column that contains the specified cell. Read-only Long.
- `Width As Single  (read/write)`  
  Returns or sets the width of a table cell, in points. Read/write Single.
- `Height As Single  (read/write)`  
  Returns or sets the height of the specified table cell. .
- `HeightRule As WdRowHeightRule  (read/write)`  
  Returns or sets a WdRowHeightRule constant that represents the rule for determining the height of the specified cells or rows. Read/write.
- `VerticalAlignment As WdCellVerticalAlignment  (read/write)`  
  Returns or sets the vertical alignment of text in one or more cells of a table. Read/write WdCellVerticalAlignment.
- `Column As Column  (read-only)`  
  Returns a Column object that represents the table column containing the specified cell. Read-only.
- `Row As Row  (read-only)`  
  Returns a Row object that represents the row containing the specified cell.
- `Next As Cell  (read-only)`  
  Returns a Cell object that represents the next table cell in the Cells collection. Read-only.
- `Previous As Cell  (read-only)`  
  Returns a Cell object that represents the previous table cell in the Cells collection. Read-only.
- `Shading As Shading  (read-only)`  
  Returns a Shading object that refers to the shading formatting for the specified object.
- `Borders As Borders  (read/write)`  
  Returns a Borders collection that represents all the borders for the specified object.
- `Tables As Tables  (read-only)`  
  Returns a Tables collection that represents all the nested tables inside the specified table cell. Read-only.
- `NestingLevel As Long  (read-only)`  
  Returns the nesting level of the specified cell. Read-only Long.
- `WordWrap As Boolean  (read/write)`  
  True if Microsoft Word wraps text to multiple lines and lengthens the cell so that the cell width remains the same. Read/write Boolean.
- `PreferredWidth As Single  (read/write)`  
  Returns or sets the preferred width (in points or as a percentage of the window width) for the specified cell. Read/write Single.
- `FitText As Boolean  (read/write)`  
  True if Microsoft Word visually reduces the size of text typed into a cell so that it fits within the column width. Read/write Boolean.
- `TopPadding As Single  (read/write)`  
  Returns or sets the amount of space (in points) to add above the contents of a single cell or all the cells in a table. Read/write Single.
- `BottomPadding As Single  (read/write)`  
  Returns or sets the amount of space (in points) to add below the contents of a single cell or all the cells in a table. Read/write Single.
- `LeftPadding As Single  (read/write)`  
  Returns or sets the amount of space (in points) to add to the left of the contents of a single cell or all the cells in a table. Read/write Single.
- `RightPadding As Single  (read/write)`  
  Returns or sets the amount of space (in points) to add to the right of the contents of a single cell or all the cells in a table. Read/write Single.
- `ID As String  (read/write)`  
  Returns or sets the identifying label for the specified object when the current document is saved as a webpage. Read/write String.
- `PreferredWidthType As WdPreferredWidthType  (read/write)`  
  Returns or sets the preferred unit of measurement to use for the width of the specified cell. Read-only WdPreferredWidthType.

## Methods (8)

- `Select()`  
  Selects the specified object.
- `Delete([ShiftCells As Variant])`  
  Deletes a table cell or cells and optionally controls how the remaining cells are shifted.
    - `ShiftCells As Variant` (optional): The direction in which the remaining cells are to be shifted. Can be any WdDeleteCells constant. If omitted, cells to the right of the last deleted cell are shifted left.
- `Formula([Formula As Variant], [NumFormat As Variant])`  
  Inserts an = (Formula) field that contains the specified formula into a table cell.
    - `Formula As Variant` (optional): The mathematical formula you want the = (Formula) field to evaluate. Spreadsheet-type references to table cells are valid. For example, "=SUM(A4:C4)" specifies the first three values in the fourth row. For more information about the = (Formula) field, see Field codes:= (Formula) field.
    - `NumFormat As Variant` (optional): A format for the result of the = (Formula) field. For information about the types of formats you can apply, see Numeric Picture (\#) field switch.
- `SetWidth(ColumnWidth As Single, RulerStyle As WdRulerStyle)`  
  Sets the width of columns or cells in a table.
    - `ColumnWidth As Single` (required): The width of the specified column or columns, in points.
    - `RulerStyle As WdRulerStyle` (required): Controls the way Word adjusts cell widths.
- `SetHeight(RowHeight As Variant, HeightRule As WdRowHeightRule)`  
  Sets the height of table cells.
    - `RowHeight As Variant` (required): The height of the row or rows, in points.
    - `HeightRule As WdRowHeightRule` (required): The rule for determining the height of the specified cells.
- `Merge(MergeTo As Cell)`  
  Merges the specified table cell with another table cell. The result is a single table cell.
    - `MergeTo As Cell` (required): The cell to be merged with.
- `Split([NumRows As Variant], [NumColumns As Variant])`  
  Splits a single table cell into multiple cells.
    - `NumRows As Variant` (optional): The number of rows that the cell or group of cells is to be split into.
    - `NumColumns As Variant` (optional): The number of columns that the cell or group of cells is to be split into.
- `AutoSum()`  
  Inserts an = (Formula) field that calculates and displays the sum of the values in table cells above or to the left of the cell specified in the expression.
