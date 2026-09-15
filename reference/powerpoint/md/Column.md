# Column

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934C5-5A91-11CF-8700-00AA0060263B}  

Represents a table column. The Column object is a member of the Columns collection. The Columns collection includes all the columns in a table.

**Example:**

```vba
ActivePresentation.Slides(2).Shapes(5).Table.Columns(1).Select
```

## Properties (4)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Cells As CellRange  (read-only)`  
  Returns a CellRange collection that represents the cells in a table column or row. Read-only.
- `Width As Single  (read/write)`  
  Returns or sets the width of the specified object, in points. Read/write.

## Methods (2)

- `Select()`  
  Selects the specified object.
- `Delete()`  
  Deletes the specified Column object.
