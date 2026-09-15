# Rows

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934C6-5A91-11CF-8700-00AA0060263B}  

A collection of Row objects that represent the rows in a table.

**Example:**

```vba
Dim i As Integer

With ActivePresentation.Slides(2).Shapes(4).Table

    For i = 1 To .Rows.Count

        .Rows.Height = 160

    Next i

End With
```

## Properties (3)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.

## Methods (2)

- `Item(Index As Long) As Row`  
  Returns a single Row object from the specified Rows collection.
    - `Index As Long` (required): The index number of the single Row object in the collection to be returned.
- `Add([BeforeRow As Long]) As Row`  
  Returns a Row object that represents a row being added to a table.
    - `BeforeRow As Long` (optional): The row before which the row is to be added.
