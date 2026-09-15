# TableStyle

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {B7564E97-0519-4C68-B400-3803E7C63242}  

Represents a single style that can be applied to a table.

**Remarks:** Use the Table property of the Styles object to return a TableStyle object. Use the Borders property to apply borders to an entire table. Use the Condition method to apply borders or shading only to specified sections of a table. This example creates a new table style and formats the table with a surrounding border. Special borders and shading are applied to the first and last rows and the last column.

## Properties (17)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified TableStyle object.
- `AllowPageBreaks As Boolean  (read/write)`  
  Allows Microsoft Word to break the specified table across pages. Read/write Boolean.
- `Borders As Borders  (read/write)`  
  Returns a Borders collection that represents all the borders for the specified object.
- `BottomPadding As Single  (read/write)`  
  Returns or sets the amount of space (in points) to add below the contents of a single cell or all the cells in a table. Read/write Single.
- `LeftPadding As Single  (read/write)`  
  Returns or sets the amount of space (in points) to add to the left of the contents of all the cells in a table. Read/write Single.
- `TopPadding As Single  (read/write)`  
  Returns or sets the amount of space (in points) to add above the contents of all the cells in a table. Read/write Single.
- `RightPadding As Single  (read/write)`  
  Returns or sets the amount of space (in points) to add to the right of the contents of all the cells in a table. Read/write Single.
- `Alignment As WdRowAlignment  (read/write)`  
  Returns or sets a WdRowAlignment constant that represents the alignment for the specified table style. Read/write.
- `Spacing As Single  (read/write)`  
  Returns or sets the spacing (in points) between the cells in a table style. Read/write Single.
- `TableDirection As WdTableDirection  (read/write)`  
  Returns or sets the direction in which Microsoft Word orders cells in the specified table style. Read/write WdTableDirection.
- `AllowBreakAcrossPage As Long  (read/write)`  
  Sets or returns a Long indicating whether lines in the rows of tables formatted with a specified style break across pages. Read/write.
- `LeftIndent As Single  (read/write)`  
  Returns or sets a Single that represents the left indent value (in points) for the rows in the specified table style. Read/write.
- `Shading As Shading  (read-only)`  
  Returns a Shading object that refers to the shading formatting for the specified table style.
- `RowStripe As Long  (read/write)`  
  Returns or sets a Long that represents the number of rows to include in the banding when a style specifies odd- or even-row banding. Read/write.
- `ColumnStripe As Long  (read/write)`  
  Returns or sets a Long that represents the number of columns in the banding when a style specifies odd- or even-column banding. Read/write.

## Methods (1)

- `Condition(ConditionCode As WdConditionCode) As ConditionalStyle`  
  Returns a ConditionalStyle object that represents special style formatting for a portion of a table.
    - `ConditionCode As WdConditionCode` (required): The area of the table to which to apply the formatting.
