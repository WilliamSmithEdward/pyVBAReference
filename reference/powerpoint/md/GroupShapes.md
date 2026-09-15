# GroupShapes

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {9149347B-5A91-11CF-8700-00AA0060263B}  

Represents the individual shapes within a grouped shape. Each shape is represented by a Shape object. Using the Item method with this object, you can work with single shapes within a group without having to ungroup them.

**Example:**

```vba
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes
    .AddShape(msoShapeIsoscelesTriangle, 10, _
        10, 100, 100).Name = "shpOne"

    .AddShape(msoShapeIsoscelesTriangle, 150, _
        10, 100, 100).Name = "shpTwo"

    .AddShape(msoShapeIsoscelesTriangle, 300, _
        10, 100, 100).Name = "shpThree"

    With .Range(Array("shpOne", "shpTwo", "shpThree")).Group
        .Fill.PresetTextured msoTextureBlueTissuePaper
        .GroupItems(2).Fill.PresetTextured msoTextureGreenMarble
    End With
End With
```

## Properties (5)

- `Application As Object  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Creator As Long  (read-only)`  
  Returns a Long that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in Microsoft PowerPoint, this property returns the hexadecimal number 50575054. Read-only.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`

## Methods (2)

- `Item(Index As Variant) As Shape`  
  Returns a single Shape object from the specified GroupShapes collection.
    - `Index As Variant` (required): The name or index number of the single Shape object in the collection to be returned.
- `Range(Index As Variant) As ShapeRange`  
  Returns a ShapeRange object.
    - `Index As Variant` (required): The individual shapes that are to be included in the range. Can be an Integer that specifies the index number of the shape, a String that specifies the name of the shape, or an array that contains either integers or strings. If this argument is omitted, the Range method returns all the objects in the specified collection.
