# Cell

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934C9-5A91-11CF-8700-00AA0060263B}  

Represents a table cell. The Cell object is a member of the CellRange collection. The CellRange collection represents all the cells in the specified column or row. To use the CellRange collection, use the Cells keyword.

**Remarks:** You cannot programmatically add cells to or delete cells from a PowerPoint table. Use the Add method of the Columns or Rows collections to add a column or row to a table. Use the Delete method of the Columns or Rows collections to delete a column or row from a table.

**Example:**

```vba
With ActivePresentation.Slides(2).Shapes(5).Table

    .Cell(1, 1).Merge MergeTo:=.Cell(1, 2)

End With
```

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Shape As Shape  (read-only)`  
  Returns a Shape object that represents a shape in a table cell. Read-only.
- `Borders As Borders  (read-only)`  
  Returns a Borders collection that represents the borders and diagonal lines for the specified Cell object or CellRange collection. Read-only.
- `Selected As Boolean  (read-only)`  
  Returns True if the specified table cell is selected. Read-only.

## Methods (3)

- `Merge(MergeTo As Cell)`  
  Merges one table cell with another. The result is a single table cell.
    - `MergeTo As Cell` (required): The Cell object to be merged with.
- `Split(NumRows As Long, NumColumns As Long)`  
  Splits a single table cell into multiple cells.
    - `NumRows As Long` (required): Number of rows that the cell is being split into.
    - `NumColumns As Long` (required): Number of columns that the cell is being split into.
- `Select()`  
  Selects the specified object.
