# Rows

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002094C-0000-0000-C000-000000000046}  

A collection of Row objects that represent the table rows in the specified selection, range, or table.

**Remarks:** Use the Rows property to return the Rows collection. The following example centers rows in the first table in the active document between the left and right margins. Use the Add method to add a row to a table. The following example inserts a row before the first row in the selection. Use Rows (Index), where Index is the index number, to return a single Row object. The index number represents the position of the row in the selection, range, or table. The following example deletes the first row in the first table in the active document.

## Properties (28)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of rows in the collection. Read-only.
- `AllowBreakAcrossPages As Long  (read/write)`  
  True if the text in a table row or rows are allowed to split across a page break. Read/write Long.
- `Alignment As WdRowAlignment  (read/write)`  
  Returns or sets a WdRowAlignment constant that represents the alignment for the specified rows. Read/write.
- `HeadingFormat As Long  (read/write)`  
  True if the specified row or rows are formatted as a table heading. Read/write Long.
- `SpaceBetweenColumns As Single  (read/write)`  
  Returns or sets the distance (in points) between text in adjacent columns of the specified row or rows. Read/write Single.
- `Height As Single  (read/write)`  
  Returns or sets the height of the specified rows in a table. Read/write Single.
- `HeightRule As WdRowHeightRule  (read/write)`  
  Returns or sets the rule for determining the height of the specified cells or rows. Read/write WdRowHeightRule.
- `LeftIndent As Single  (read/write)`  
  Returns or sets a Single that represents the left indent value (in points) for the specified table rows. Read/write.
- `First As Row  (read-only)`  
  Returns a Row object that represents the first item in the Rows collection.
- `Last As Row  (read-only)`  
  Returns the last item in the Rows collection as a Row object.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Rows object.
- `Borders As Borders  (read/write)`  
  Returns a Borders collection that represents all the borders for the specified object.
- `Shading As Shading  (read-only)`  
  Returns a Shading object that refers to the shading formatting for the specified object.
- `WrapAroundText As Long  (read/write)`  
  Returns or sets whether text should wrap around the specified rows. Read/write Long.
- `DistanceTop As Single  (read/write)`  
  Returns or sets the distance (in points) between the document text and the top edge of the specified table. Read/write Single.
- `DistanceBottom As Single  (read/write)`  
  Returns or sets the distance (in points) between the document text and the bottom edge of the specified table. Read/write Single.
- `DistanceLeft As Single  (read/write)`  
  Returns or sets the distance (in points) between the document text and the left edge of the specified table. Read/write Single.
- `DistanceRight As Single  (read/write)`  
  Returns or sets the distance (in points) between the document text and the right edge of the specified table. Read/write Single.
- `HorizontalPosition As Single  (read/write)`  
  Returns or sets the horizontal distance between the edge of the rows and the item specified by the RelativeHorizontalPosition property. Read/write Single.
- `VerticalPosition As Single  (read/write)`  
  Returns or sets the vertical distance between the edge of the rows and the item specified by the RelativeVerticalPosition property. Read/write Single.
- `RelativeHorizontalPosition As WdRelativeHorizontalPosition  (read/write)`  
  Specifies the relative horizontal position of a group of rows. Read/write WdRelativeHorizontalPosition.
- `RelativeVerticalPosition As WdRelativeVerticalPosition  (read/write)`  
  Specifies the relative vertical position of a group of rows. Read/write WdRelativeVerticalPosition.
- `AllowOverlap As Long  (read/write)`  
  Returns or sets a value that specifies whether the specified rows can overlap other rows.
- `NestingLevel As Long  (read-only)`  
  Returns the nesting level of the specified table rows. Read-only Long.
- `TableDirection As WdTableDirection  (read/write)`  
  Returns or sets the direction in which Microsoft Word orders cells in the specified table or row. Read/write WdTableDirection.

## Methods (8)

- `Item(Index As Long) As Row`  
  Returns an individual Row object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
- `Add([BeforeRow As Variant]) As Row`  
  Returns a Row object that represents a row added to a table.
    - `BeforeRow As Variant` (optional): A Row object that represents the row that will appear immediately below the new row.
- `Select()`  
  Selects a collection of rows in a table.
- `Delete()`  
  Deletes the specified table rows.
- `SetLeftIndent(LeftIndent As Single, RulerStyle As WdRulerStyle)`  
  Sets the indentation for a row or rows in a table.
    - `LeftIndent As Single` (required): The distance (in points) between the current left edge of the specified row or rows and the desired left edge.
    - `RulerStyle As WdRulerStyle` (required): Controls the way Microsoft Word adjusts the table when the left indent is changed. The WdRulerStyle behavior applies to left-aligned tables. The WdRulerStyle behavior for center- and right-aligned tables can be unexpected; in these cases, use the SetLeftIndent method with care.
- `SetHeight(RowHeight As Single, HeightRule As WdRowHeightRule)`  
  Sets the height of table rows.
    - `RowHeight As Single` (required): The height of the row or rows, in points.
    - `HeightRule As WdRowHeightRule` (required): The rule for determining the height of the specified rows.
- `DistributeHeight()`  
  Adjusts the height of the specified rows or cells so that they're equal.
- `ConvertToText([Separator As Variant], [NestedTables As Variant]) As Range`  
  Converts rows in a table to text and returns a Range object that represents the delimited text.
    - `Separator As Variant` (optional): The character that delimits the converted columns (paragraph marks delimit the converted rows). Can be any following WdTableFieldSeparator constants: wdSeparateByCommas, wdSeparateByDefaultListSeparator, wdSeparateByParagraphs, or wdSeparateByTabs (Default).
    - `NestedTables As Variant` (optional): True if nested tables are converted to text. This argument is ignored if Separator is not wdSeparateByParagraphs. The default value is True.
