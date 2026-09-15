# FreeformBuilder

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493478-5A91-11CF-8700-00AA0060263B}  

Represents the geometry of a freeform while it is being built.

**Example:**

```vba
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.BuildFreeform(msoEditingCorner, 360, 200)
    .AddNodes msoSegmentCurve, msoEditingCorner, _
        380, 230, 400, 250, 450, 300
    .AddNodes msoSegmentCurve, msoEditingAuto, 480, 200
    .AddNodes msoSegmentLine, msoEditingAuto, 480, 40
    .AddNodes msoSegmentLine, msoEditingAuto, 360, 200
    .ConvertToShape
End With
```

## Properties (3)

- `Application As Object  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Creator As Long  (read-only)`  
  Returns a Long that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in Microsoft PowerPoint, this property returns the hexadecimal number 50575054. Read-only.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.

## Methods (2)

- `AddNodes(SegmentType As MsoSegmentType, EditingType As MsoEditingType, X1 As Single, Y1 As Single, [X2 As Single], [Y2 As Single], [X3 As Single], [Y3 As Single])`  
  Inserts a new segment at the end of the freeform that's being created, and adds the nodes that define the segment. Use this method as many times as you want to add nodes to the freeform you are creating. When you finish adding nodes, use the ConvertToShape method to create the freeform you've just defined. To add nodes to a freeform after it is been created, use the Insert method of the ShapeNodes collection.
    - `SegmentType As MsoSegmentType` (required): The type of segment to be added.
    - `EditingType As MsoEditingType` (required): The editing property of the vertex. If SegmentType is msoSegmentLine, EditingType must be msoEditingAuto.
    - `X1 As Single` (required): If the EditingType of the new segment is msoEditingAuto, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the endpoint of the new segment. If the EditingType of the new node is msoEditingCorner, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the first control point for the new segment.
    - `Y1 As Single` (required): If the EditingType of the new segment is msoEditingAuto, this argument specifies the vertical distance (in points) from the upper-left corner of the document to the endpoint of the new segment. If the EditingType of the new node is msoEditingCorner, this argument specifies the vertical distance (in points) from the upper-left corner of the document to the first control point for the new segment.
    - `X2 As Single` (optional): If the EditingType of the new segment is msoEditingCorner, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the second control point for the new segment. If the EditingType of the new segment is msoEditingAuto, don't specify a value for this argument.
    - `Y2 As Single` (optional): If the EditingType of the new segment is msoEditingCorner, this argument specifies the vertical distance (in points) from the upper-left corner of the document to the second control point for the new segment. If the EditingType of the new segment is msoEditingAuto, don't specify a value for this argument.
    - `X3 As Single` (optional): If the EditingType of the new segment is msoEditingCorner, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the endpoint of the new segment. If the EditingType of the new segment is msoEditingAuto, don't specify a value for this argument.
    - `Y3 As Single` (optional): If the EditingType of the new segment is msoEditingCorner, this argument specifies the vertical distance (in points) from the upper-left corner of the document to the endpoint of the new segment. If the EditingType of the new segment is msoEditingAuto, don't specify a value for this argument.
- `ConvertToShape() As Shape`  
  Creates a shape that has the geometric characteristics of the specified FreeformBuilder object. Returns a Shape object that represents the new shape.
