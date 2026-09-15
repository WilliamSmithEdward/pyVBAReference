# ConditionalStyle

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {1498F56D-ED33-41F9-B37B-EF30E50B08AC}  

Represents special formatting applied to specified areas of a table when the selected table is formatted with a specified table style.

**Remarks:** Use the Condition method of the TableStyle object to return a ConditionalStyle object. The Shading property can be used to apply shading to specified areas of a table. This example selects the first table in the active document and applies shading to alternate rows and columns. This example assumes that there is a table in the active document and that it is formatted using the Table Grid style. Use the Borders property to apply borders to specified areas of a table. This example selects the first table in the active document and applies borders to the first and last row and first column. This example assumes that there is a table in the active document and that it is formatted using the Table Grid style.

## Properties (11)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ConditionalStyle object.
- `Shading As Shading  (read-only)`  
  Returns a Shading object that represents the shading formatting for the style formatting of a table.
- `Borders As Borders  (read/write)`  
  Returns a Borders collection that represents all the borders for the specified ConditionalStyle object.
- `BottomPadding As Single  (read/write)`  
  Returns or sets a Single that represents the amount of space (in points) to add below the contents of a single cell or all the cells in a table. Read/write.
- `TopPadding As Single  (read/write)`  
  Returns or sets the amount of space (in points) to add above the contents of a single cell or all the cells in a table. Read/write Single.
- `LeftPadding As Single  (read/write)`  
  Returns or sets the amount of space (in points) to add to the left of the contents of a single cell or all the cells in a table. Read/write Single.
- `RightPadding As Single  (read/write)`  
  Returns or sets the amount of space (in points) to add to the right of the contents of a single cell or all the cells in a table. Read/write Single.
- `ParagraphFormat As ParagraphFormat  (read/write)`  
  Returns or sets a ParagraphFormat object that represents the paragraph settings for the specified conditional style. Read/write.
- `Font As Font  (read/write)`  
  Returns or sets a Font object that represents the character formatting of the specified object. Read/write.
