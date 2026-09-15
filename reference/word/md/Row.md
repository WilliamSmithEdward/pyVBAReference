# Row

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020950-0000-0000-C000-000000000046}  

Represents a row in a table. The Row object is a member of the Rows collection. The Rows collection includes all the rows in the specified selection, range, or table.

**Remarks:** Use Rows (Index), where Index is the index number, to return a single Row object. The index number represents the position of the row in the selection, range, or table. The following example deletes the first row in the first table in the active document. Use the Add method to add a row to a table. The following example inserts a row before the first row in the selection. Use the Cells property to modify the individual cells in a Row object. The following example adds a table to the selection and then inserts numbers into each cell in the second row of the table.

## Properties (21)

- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that is contained within the specified table row.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Row object.
- `AllowBreakAcrossPages As Long  (read/write)`  
  True if the text in a table row or rows are allowed to split across a page break. Read/write Long.
- `Alignment As WdRowAlignment  (read/write)`  
  Returns or sets a WdRowAlignment constant that represents the alignment for the specified rows. Read/write.
- `HeadingFormat As Long  (read/write)`  
  True if the specified row or rows are formatted as a table heading. Rows formatted as table headings are repeated when a table spans more than one page. Can be True, False or wdUndefined. Read/write Long.
- `SpaceBetweenColumns As Single  (read/write)`  
  Returns or sets the distance (in points) between text in adjacent columns of the specified row or rows. Read/write Single.
- `Height As Single  (read/write)`  
  Returns or sets the height (in points) of the specified row in a table. Read/write Single.
- `HeightRule As WdRowHeightRule  (read/write)`  
  Returns or sets the rule for determining the height of the specified cells or rows. Read/write WdRowHeightRule.
- `LeftIndent As Single  (read/write)`  
  Returns or sets a Single that represents the left indent value (in points) for the specified table row. Read/write.
- `IsLast As Boolean  (read-only)`  
  True if the specified row is the last one in the table. Read-only Boolean.
- `IsFirst As Boolean  (read-only)`  
  True if the specified row is the first one in the table. Read-only Boolean.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
- `Cells As Cells  (read-only)`  
  Returns a Cells collection that represents the table cells in a column, row, selection, or range. Read-only.
- `Borders As Borders  (read/write)`  
  Returns a Borders collection that represents all the borders for the specified object.
- `Shading As Shading  (read-only)`  
  Returns a Shading object that refers to the shading formatting for the specified object.
- `Next As Row  (read-only)`  
  Returns a Row object that represents the table row that is next in the collection of rows in a table. Read-only.
- `Previous As Row  (read-only)`  
  Returns a Row object that represents the table row that is previous to the specified row. Read-only.
- `NestingLevel As Long  (read-only)`  
  Returns the nesting level of the specified table row. Read-only Long.
- `ID As String  (read/write)`  
  Returns or sets the identifying label for the specified table row when the document is saved as a webpage. Read/write String.

## Methods (5)

- `Select()`  
  Selects the specified table row.
- `Delete()`  
  Deletes the specified table row.
- `SetLeftIndent(LeftIndent As Single, RulerStyle As WdRulerStyle)`  
  Sets the indentation for a row in a table.
    - `LeftIndent As Single` (required): The distance (in points) between the current left edge of the specified row or rows and the desired left edge.
    - `RulerStyle As WdRulerStyle` (required): Controls the way Word adjusts the table when the left indent is changed.
- `SetHeight(RowHeight As Single, HeightRule As WdRowHeightRule)`  
  Sets the height of a table row.
    - `RowHeight As Single` (required): The height of the row, in points.
    - `HeightRule As WdRowHeightRule` (required): The rule for determining the height of the specified rows.
- `ConvertToText([Separator As Variant], [NestedTables As Variant]) As Range`  
  Converts a table to text and returns a Range object that represents the delimited text.
    - `Separator As Variant` (optional): The character that delimits the converted columns (paragraph marks delimit the converted rows). Can be any following WdTableFieldSeparator constants: wdSeparateByCommas, wdSeparateByDefaultListSeparator, wdSeparateByParagraphs, or wdSeparateByTabs (Default).
    - `NestedTables As Variant` (optional): True if nested tables are converted to text. This argument is ignored if Separator is not wdSeparateByParagraphs. The default value is True.
