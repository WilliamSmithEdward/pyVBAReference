# TabStops

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493493-5A91-11CF-8700-00AA0060263B}  

A collection of all the TabStop objects on one ruler.

**Example:**

```vba
With ActivePresentation.Slides(1).Shapes(2) _
        .TextFrame.Ruler.TabStops
    For t = .Count To 1 Step -1
        .Item(t).Clear
    Next
End With
```

## Properties (4)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `DefaultSpacing As Single  (read/write)`  
  Returns or sets the default tab-stop spacing for the specified text, in points. Read/write.

## Methods (2)

- `Item(Index As Long) As TabStop`  
  Returns a single tab stop from the specified TabStops collection.
    - `Index As Long` (required): The index number of the single TabStop object in the collection to be returned.
- `Add(Type As PpTabStopType, Position As Single) As TabStop`  
  Creates a tab stop and adds it to the TabStops collection.
    - `Type As PpTabStopType` (required): The type of the tab stop to be added.
    - `Position As Single` (required): The position of the tab stop in the tab stops collection.
