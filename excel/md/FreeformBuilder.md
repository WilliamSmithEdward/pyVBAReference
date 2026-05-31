# FreeformBuilder

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002443F-0000-0000-C000-000000000046}  

Represents the geometry of a freeform while it's being built.

**Remarks:** Use the BuildFreeform method of the Shapes object to return a FreeformBuilder object. Use the AddNodes method to add nodes to the freeform. Use the ConvertToShape method to create the shape defined in the FreeformBuilder object and add it to the Shapes collection.

**Example:**

```vba
Set myDocument = Worksheets(1)
With myDocument.Shapes.BuildFreeform(msoEditingCorner, 360, 200)
 .AddNodes msoSegmentCurve, msoEditingCorner, _
 380, 230, 400, 250, 450, 300
 .AddNodes msoSegmentCurve, msoEditingAuto, 480, 200
 .AddNodes msoSegmentLine, msoEditingAuto, 480, 400
 .AddNodes msoSegmentLine, msoEditingAuto, 360, 200
 .ConvertToShape
End With
```

## Properties (3)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.

## Methods (2)

- `AddNodes(SegmentType As MsoSegmentType, EditingType As MsoEditingType, X1 As Single, Y1 As Single, [X2 As Variant], [Y2 As Variant], [X3 As Variant], [Y3 As Variant])`  
  Adds a point in the current shape, and then draws a line from the current node to the last node that was added.
    - `SegmentType As MsoSegmentType` (required): The type of segment to be added.
    - `EditingType As MsoEditingType` (required): The editing property of the vertex.
    - `X1 As Single` (required): If the _EditingType_ of the new segment is msoEditingAuto, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the end point of the new segment. If the _EditingType_ of the new node is msoEditingCorner, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the first control point for the new segment.
    - `Y1 As Single` (required): If the _EditingType_ of the new segment is msoEditingAuto, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the end point of the new segment. If the _EditingType_ of the new node is msoEditingCorner, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the first control point for the new segment.
    - `X2 As Variant` (optional): If the _EditingType_ of the new segment is msoEditingCorner, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the second control point for the new segment. If the _EditingType_ of the new segment is msoEditingAuto, don't specify a value for this argument.
    - `Y2 As Variant` (optional): If the _EditingType_ of the new segment is msoEditingCorner, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the second control point for the new segment. If the _EditingType_ of the new segment is msoEditingAuto, don't specify a value for this argument.
    - `X3 As Variant` (optional): If the _EditingType_ of the new segment is msoEditingCorner, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the second control point for the new segment. If the _EditingType_ of the new segment is msoEditingAuto, don't specify a value for this argument.
    - `Y3 As Variant` (optional): If the _EditingType_ of the new segment is msoEditingCorner, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the second control point for the new segment. If the _EditingType_ of the new segment is msoEditingAuto, don't specify a value for this argument.
- `ConvertToShape() As Shape`  
  Creates a shape that has the geometric characteristics of the specified FreeformBuilder object. Returns a Shape object that represents the new shape.
