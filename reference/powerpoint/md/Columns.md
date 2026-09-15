# Columns

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934C4-5A91-11CF-8700-00AA0060263B}  

A collection of Column objects that represent the columns in a table.

**Example:**

```vba
Dim ColCount, sl, sh As Integer

With ActivePresentation
    For sl = 1 To .Slides.Count
        For sh = 1 To .Slides(sl).Shapes.Count
            If .Slides(sl).Shapes(sh).HasTable Then
                ColCount = .Slides(sl).Shapes(sh) _
                    .Table.Columns.Count
                MsgBox "Shape " & sh & " on slide " & sl & _
                    " contains the first table and has " & _
                    ColCount & " columns."
                Exit Sub
            End If
        Next
    Next
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

- `Item(Index As Long) As Column`  
  Returns a single Column object from the specified Columns collection.
    - `Index As Long` (required): The index number of the single Column object in the collection to be returned.
- `Add([BeforeColumn As Long]) As Column`  
  Adds a new column to an existing table. Returns a Column object that represents the new table column.
    - `BeforeColumn As Long` (optional): The index number that specifies the table column before which the new column will be inserted.
