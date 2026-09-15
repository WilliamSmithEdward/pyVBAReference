# FreeformBuilder

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209C9-0000-0000-C000-000000000046}  

Represents the geometry of a freeform while it is being built.

**Remarks:** Use the BuildFreeform method of the Shapes or CanvasShapes object to return a FreeformBuilder object. Use the AddNodes method to add nodes to the freeform. Use the ConvertToShape method to create the shape defined in the FreeformBuilder object and add it to the Shapes collection. The following example adds a freeform with four segments to the active document.

## Properties (3)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified FreeformBuilder object.

## Methods (2)

- `AddNodes(SegmentType As MsoSegmentType, EditingType As MsoEditingType, X1 As Single, Y1 As Single, [X2 As Single], [Y2 As Single], [X3 As Single], [Y3 As Single])`  
  Inserts a new segment at the end of the freeform that's being created, and adds the nodes that define the segment.
    - `SegmentType As MsoSegmentType` (required): The type of segment to be added.
    - `EditingType As MsoEditingType` (required): The editing property of the vertex. If SegmentType is msoSegmentLine, EditingType must be msoEditingAuto.
    - `X1 As Single` (required): If the EditingType of the new segment is msoEditingAuto, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the endpoint of the new segment. If the EditingType of the new node is msoEditingCorner, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the first control point for the new segment.
    - `Y1 As Single` (required): If the EditingType of the new segment is msoEditingAuto, this argument specifies the vertical distance (in points) from the upper-left corner of the document to the endpoint of the new segment. If the EditingType of the new node is msoEditingCorner, this argument specifies the vertical distance (in points) from the upper-left corner of the document to the first control point for the new segment.
    - `X2 As Single` (optional): If the EditingType of the new segment is msoEditingCorner, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the second control point for the new segment. If the EditingType of the new segment is msoEditingAuto, don't specify a value for this argument.
    - `Y2 As Single` (optional): If the EditingType of the new segment is msoEditingCorner, this argument specifies the vertical distance (in points) from the upper-left corner of the document to the second control point for the new segment. If the EditingType of the new segment is msoEditingAuto, don't specify a value for this argument.
    - `X3 As Single` (optional): If the EditingType of the new segment is msoEditingCorner, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the endpoint of the new segment. If the EditingType of the new segment is msoEditingAuto, don't specify a value for this argument.
    - `Y3 As Single` (optional): If the EditingType of the new segment is msoEditingCorner, this argument specifies the vertical distance (in points) from the upper-left corner of the document to the endpoint of the new segment. If the EditingType of the new segment is msoEditingAuto, don't specify a value for this argument.
- `ConvertToShape([Anchor As Variant]) As Shape`  
  Creates a shape that has the geometric characteristics of the specified object. Returns a Shape object that represents the new shape.
    - `Anchor As Variant` (optional): A Range object that represents the text to which the shape is bound. If Anchor is specified, the anchor is positioned at the beginning of the first paragraph in the anchoring range. If this argument is omitted, the anchoring range is selected automatically and the shape is positioned relative to the top and left edges of the page.
