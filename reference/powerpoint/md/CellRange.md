# CellRange

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934C8-5A91-11CF-8700-00AA0060263B}  

A collection of Cell objects in a table column or row. The CellRange collection represents all the cells in the specified column or row. To use the CellRange collection, use the Cells keyword.

**Remarks:** Although the collection object is named CellRange and is shown in the Object Browser, this keyword is not used in programming the PowerPoint object model. The keyword Cells is used instead. You cannot programmatically add cells to or delete cells from a PowerPoint table. Use the AddTable method with the Table object to add a new table. Use the Add method of the Columns or Rows collections to add a column or row to a table. Use the Delete method of the Columns or Rows collections to delete a column or row from a table.

**Example:**

```vba
With ActivePresentation.Slides(2).Shapes(5).Table.Columns(1).Cells

    .Borders(ppBorderRight).DashStyle = msoLineDash

End With
```

## Properties (4)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Borders As Borders  (read-only)`  
  Returns a Borders collection that represents the borders and diagonal lines for the specified Cell object or CellRange collection. Read-only.

## Methods (1)

- `Item(Index As Long) As Cell`  
  Returns a single Cell object from the specified CellRange collection.
    - `Index As Long` (required): The index number of the single Cell object in the collection to be returned.
