# SlideShowWindows

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493456-5A91-11CF-8700-00AA0060263B}  

A collection of all the SlideShowWindow objects that represent the open slide shows in Microsoft PowerPoint.

**Example:**

```vba
With SlideShowWindows(1)

    If .IsFullScreen Then

        .Height = .Height - 20

    End If

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

- `Item(Index As Long) As SlideShowWindow`  
  Returns a single SlideShowWindow object from the specified SlideShowWindows collection.
    - `Index As Long` (required): The index number of the single SlideShowWindow object in the collection to be returned.
