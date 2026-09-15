# Row

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934C7-5A91-11CF-8700-00AA0060263B}  

Represents a row in a table. The Row object is a member of the Rows collection. The Rows collection includes all the rows in the specified table.

**Example:**

```vba
ActivePresentation.Slides(2).Shapes(5).Table.Rows(1).Delete
```

## Properties (4)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Cells As CellRange  (read-only)`  
  Returns a CellRange collection that represents the cells in a table column or row. Read-only.
- `Height As Single  (read/write)`  
  Returns or sets the height of the specified object, in points. Read/write.

## Methods (2)

- `Select()`  
  Selects the specified object.
- `Delete()`  
  Deletes the specified Row object.
