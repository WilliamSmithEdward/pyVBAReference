# ObjectVerbs

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {9149348A-5A91-11CF-8700-00AA0060263B}  

Represents the collection of OLE verbs for the specified OLE object. OLE verbs are the operations supported by an OLE object. Commonly used OLE verbs are "play" and "edit."

**Example:**

```vba
With ActivePresentation.Slides(2).Shapes(1).OLEFormat

    For Each v In .ObjectVerbs

        MsgBox v

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

## Methods (1)

- `Item(Index As Long) As String`  
  Returns a single object from the specified ObjectVerbs collection.
    - `Index As Long` (required): The index number of the single object in the collection to be returned.
