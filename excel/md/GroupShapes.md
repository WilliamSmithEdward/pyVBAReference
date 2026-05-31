# GroupShapes

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002443C-0000-0000-C000-000000000046}  

Represents the individual shapes within a grouped shape.

**Remarks:** Each shape is represented by a Shape object. Using the Item method with this object, you can work with single shapes within a group without having to ungroup them.

**Example:**

```vba
Set myDocument = Worksheets(1)
With myDocument.Shapes
 .AddShape(msoShapeIsoscelesTriangle, _
 10, 10, 100, 100).Name = "shpOne"
 .AddShape(msoShapeIsoscelesTriangle, _
 150, 10, 100, 100).Name = "shpTwo"
 .AddShape(msoShapeIsoscelesTriangle, _
 300, 10, 100, 100).Name = "shpThree"
 With .Range(Array("shpOne", "shpTwo", "shpThree")).Group
 .Fill.PresetTextured msoTextureBlueTissuePaper
 .GroupItems(2).Fill.PresetTextured msoTextureGreenMarble
 End With
End With
```

## Properties (6)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.
- `_NewEnum As IUnknown  (read-only)`
- `Range As ShapeRange  (read-only)`  
  Returns a ShapeRange object that represents a subset of the shapes in a Shapes collection.

## Methods (2)

- `Item(Index As Variant) As Shape`  
  Returns a single object from a collection.
    - `Index As Variant` (required): The name or index number for the object.
- `_Default(Index As Variant) As Shape`
