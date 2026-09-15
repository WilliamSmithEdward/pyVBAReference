# ShapeNodes

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493486-5A91-11CF-8700-00AA0060263B}  

A collection of all the ShapeNode objects in the specified freeform.

**Remarks:** Each ShapeNode object represents either a node between segments in a freeform or a control point for a curved segment of a freeform. You can create a freeform manually or by using the BuildFreeformand ConvertToShapemethods.

**Example:**

```vba
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes(3).Nodes.Delete 4
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

## Methods (6)

- `Item(Index As Variant) As ShapeNode`  
  Returns a single ShapeNode object from the specified ShapeNodes collection.
    - `Index As Variant` (required): The name or index number of the single ShapeNode object in the collection to be returned.
- `Delete(Index As Long)`  
  Deletes a shape node.
    - `Index As Long` (required): Specifies the node to be deleted.
- `Insert(Index As Long, SegmentType As MsoSegmentType, EditingType As MsoEditingType, X1 As Single, Y1 As Single, [X2 As Single], [Y2 As Single], [X3 As Single], [Y3 As Single])`  
  Inserts a new segment after the specified node of the freeform.
    - `Index As Long` (required): The node that the new node is to be inserted after.
    - `SegmentType As MsoSegmentType` (required): The type of segment to be added.
    - `EditingType As MsoEditingType` (required): The editing property of the vertex.
    - `X1 As Single` (required): If the EditingType of the new segment is msoEditingAuto, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the endpoint of the new segment. If the EditingType of the new node is msoEditingCorner, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the first control point for the new segment.
    - `Y1 As Single` (required): If the EditingType of the new segment is msoEditingAuto, this argument specifies the vertical distance (in points) from the upper-left corner of the document to the endpoint of the new segment. If the EditingType of the new node is msoEditingCorner, this argument specifies the vertical distance (in points) from the upper-left corner of the document to the first control point for the new segment.
    - `X2 As Single` (optional): If the EditingType of the new segment is msoEditingCorner, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the second control point for the new segment. If the EditingType of the new segment is msoEditingAuto, don't specify a value for this argument.
    - `Y2 As Single` (optional): If the EditingType of the new segment is msoEditingCorner, this argument specifies the vertical distance (in points) from the upper-left corner of the document to the second control point for the new segment. If the EditingType of the new segment is msoEditingAuto, don't specify a value for this argument.
    - `X3 As Single` (optional): If the EditingType of the new segment is msoEditingCorner, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the endpoint of the new segment. If the EditingType of the new segment is msoEditingAuto, don't specify a value for this argument.
    - `Y3 As Single` (optional): If the EditingType of the new segment is msoEditingCorner, this argument specifies the vertical distance (in points) from the upper-left corner of the document to the endpoint of the new segment. If the EditingType of the new segment is msoEditingAuto, don't specify a value for this argument.
- `SetEditingType(Index As Long, EditingType As MsoEditingType)`  
  Sets the editing type of the specified node.
    - `Index As Long` (required): The node whose editing type is to be set.
    - `EditingType As MsoEditingType` (required): The editing type.
- `SetPosition(Index As Long, X1 As Single, Y1 As Single)`  
  Sets the location of the node specified by Index. Note that, depending on the editing type of the node, this method may affect the position of adjacent nodes.
    - `Index As Long` (required): The node whose position is to be set.
    - `X1 As Single` (required): The x-position (in points) of the new node relative to the upper-left corner of the document.
    - `Y1 As Single` (required): The y-position (in points) of the new node relative to the upper-left corner of the document.
- `SetSegmentType(Index As Long, SegmentType As MsoSegmentType)`  
  Sets the segment type of the segment that follows the specified node.
    - `Index As Long` (required): The node whose segment type is to be set.
    - `SegmentType As MsoSegmentType` (required): Specifies if the segment is straight or curved.
