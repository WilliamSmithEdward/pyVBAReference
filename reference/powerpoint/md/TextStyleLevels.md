# TextStyleLevels

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {9149349A-5A91-11CF-8700-00AA0060263B}  

A collection of all the outline text levels. This collection always contains five members, each of which is represented by a TextStyleLevel object.

**Example:**

```vba
With ActivePresentation.SlideMaster _
        .TextStyles(ppBodyStyle).Levels(1)
    With .Font
        .Name = "Arial"
        .Size = 36
    End With
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

- `Item(Level As Long) As TextStyleLevel`  
  Returns a single text style level from the specified TextStyleLevels collection.
    - `Level As Long` (required): The index number of the text style level in the collection to be returned.
