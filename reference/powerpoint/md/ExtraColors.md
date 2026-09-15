# ExtraColors

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493468-5A91-11CF-8700-00AA0060263B}  

Represents the extra colors in a presentation. The object can contain up to eight colors, each of which is represented by an red-green-blue (RGB) value.

**Example:**

```vba
With ActivePresentation
    Set rect = .Slides(1).Shapes _
        .AddShape(msoShapeRectangle, 50, 50, 100, 200)
    rect.Fill.ForeColor.RGB = .ExtraColors(1)
End With
```

## Properties (3)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.

## Methods (3)

- `Item(Index As Long) As MsoRGBType`  
  Returns a single color from the specified ExtraColors collection.
    - `Index As Long` (required): The index number of the single object in the collection to be returned.
- `Add(Type As MsoRGBType)`  
  Adds a color to the extra colors available to a presentation, if the color hasn't already been added.
    - `Type As MsoRGBType` (required): The red-green-blue (RGB) value of the color to be added.
- `Clear()`  
  Clears the extra colors in a presentation.
