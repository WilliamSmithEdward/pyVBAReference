# Table

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934C3-5A91-11CF-8700-00AA0060263B}  

Represents a table shape on a slide. The Table object is a member of the Shapes collection. The Table object contains the Columns collection and the Rows collection.

**Example:**

```vba
With ActivePresentation.Slides(2).Shapes

    For i = 1 To .Count

        If .Item(i).HasTable Then

            .Item(i).ActionSettings(ppMouseClick) _

                .Action = ppActionNextSlide

        End If

    Next

End With
```

## Properties (15)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Columns As Columns  (read-only)`  
  Returns a Columns collection that represents all the columns in a table. Read-only.
- `Rows As Rows  (read-only)`  
  Returns a Rows collection that represents all the rows in a table. Read-only.
- `TableDirection As PpDirection  (read/write)`  
  Returns or sets the direction in which the table cells are ordered. Read/write.
- `FirstRow As Boolean  (read/write)`  
  Determine whether to display special formatting for the first row of the specified table. Read/write.
- `LastRow As Boolean  (read/write)`  
  Determines whether to display special formatting for the last row of the specified table. Read/write.
- `FirstCol As Boolean  (read/write)`  
  Determines whether to display special formatting for the first column of the specified table. Read/write.
- `LastCol As Boolean  (read/write)`  
  Determines whether to display special formatting for the last column of the specified table. Read/write.
- `HorizBanding As Boolean  (read/write)`  
  Determines whether to display banded rows, in which even rows are formatted differently from odd rows. Read/write.
- `VertBanding As Boolean  (read/write)`  
  Determines whether to display banded columns, in which even columns are formatted differently from odd columns. Read/write.
- `Style As TableStyle  (read-only)`  
  Returns a TableStyle object that contains information about the specified table's current table style. Read-only.
- `Background As TableBackground  (read-only)`  
  Returns the TableBackground object associated with the specified table. Read-only.
- `AlternativeText As String  (read/write)`  
  Returns or sets the alternative text associated with a shape in a Web presentation. Read/write.
- `Title As String  (read/write)`  
  Returns or sets the title of the Table object. Read/write.

## Methods (3)

- `Cell(Row As Long, Column As Long) As Cell`  
  Returns a Cell object that represents a cell in a table.
    - `Row As Long` (required): The number of the row in the table to return. Can be an integer between 1 and the number of rows in the table.
    - `Column As Long` (required): The number of the column in the table to return. Can be an integer between 1 and the number of columns in the table.
- `ScaleProportionally(scale As Single)`  
  Scales all cell heights and widths, font sizes, and internal margins in the table by a specified proportion.
    - `scale As Single` (required): The proportion to scale the table, between 0.01 and 100. For example, a scale value of 1 keeps the table layout unchanged; a value of 2 makes it twice as large; a value of 0.5 makes it half the size.
- `ApplyStyle([StyleID As String], [SaveFormatting As Boolean])`  
  Applies a table style to the specified table.
    - `StyleID As String` (optional): The identifier of the table style to apply.
    - `SaveFormatting As Boolean` (optional): True preserves table formatting.
