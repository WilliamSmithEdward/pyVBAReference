# ShapeNodes

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209CE-0000-0000-C000-000000000046}  

A collection of all the ShapeNode objects in the specified freeform. Each ShapeNode object represents either a node between segments in a freeform or a control point for a curved segment of a freeform.

**Remarks:** You can create a freeform manually or by using the BuildFreeform and ConvertToShape methods. Use the Nodes property to return the ShapeNodes collection. The following example deletes node four in shape three on the active document. For this example to work, shape three must be a freeform with at least four nodes. Use the Insert method to create a new node and add it to the ShapeNodes collection. The following example adds a smooth node with a curved segment after node four in shape three on the active document. For this example to work, shape three must be a freeform with at least four nodes. Use Nodes (Index), where Index is the node index number, to return a single ShapeNode object. If node one in shape three on the active document is a corner point, the following example makes it a smooth point. For this example to work, shape three must be a freeform.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ShapeNodes object.
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of shape nodes in the collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`

## Methods (6)

- `Delete(Index As Long)`  
  Deletes the specified shape node.
    - `Index As Long` (required): The number within the collection of shape nodes of the shape node to delete.
- `Item(Index As Variant) As ShapeNode`  
  Returns an individual ShapeNode object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
- `SetEditingType(Index As Long, EditingType As MsoEditingType)`  
  Sets the editing type of the node specified by Index. .
    - `Index As Long` (required): The node whose editing type is to be set.
    - `EditingType As MsoEditingType` (required): The editing property of the vertex.
- `SetPosition(Index As Long, X1 As Single, Y1 As Single)`  
  Sets the location of the node specified by Index.
    - `Index As Long` (required): The node whose position is to be set.
    - `X1 As Single` (required): The position (in points) of the new node relative to the upper-left corner of the document.
- `SetSegmentType(Index As Long, SegmentType As MsoSegmentType)`  
  Sets the segment type of the segment that follows the node specified by Index.
    - `Index As Long` (required): The node whose segment type is to be set.
    - `SegmentType As MsoSegmentType` (required): Specifies if the segment is straight or curved.
- `Insert(Index As Long, SegmentType As MsoSegmentType, EditingType As MsoEditingType, X1 As Single, Y1 As Single, [X2 As Single], [Y2 As Single], [X3 As Single], [Y3 As Single])`  
  Inserts a node into a freeform shape.
    - `Index As Long` (required): The number of the shape node after which to insert a new node.
    - `SegmentType As MsoSegmentType` (required): The type of line that connects the inserted node to the neighboring nodes.
    - `EditingType As MsoEditingType` (required): The editing property of the inserted node.
    - `X1 As Single` (required): If the EditingType of the new segment is msoEditingAuto, this argument specifies the horizontal distance, measured in points, from the upper-left corner of the document to the starting point of the new segment. If the EditingType of the new node is msoEditingCorner, this argument specifies the horizontal distance, measured in points, from the upper-left corner of the document to the first control point for the new segment.
    - `Y1 As Single` (required): If the EditingType of the new segment is msoEditingAuto, this argument specifies the vertical distance, measured in points, from the upper-left corner of the document to the starting point of the new segment. If the EditingType of the new node is msoEditingCorner, this argument specifies the vertical distance, measured in points, from the upper-left corner of the document to the first control point for the new segment.
    - `X2 As Single` (optional): If the EditingType of the new segment is msoEditingCorner, this argument specifies the horizontal distance, measured in points, from the upper-left corner of the document to the second control point for the new segment. If the EditingType of the new segment is msoEditingAuto, don't specify a value for this argument.
    - `Y2 As Single` (optional): If the EditingType of the new segment is msoEditingCorner, this argument specifies the vertical distance, measured in points, from the upper-left corner of the document to the second control point for the new segment. If the EditingType of the new segment is msoEditingAuto, don't specify a value for this argument.
    - `X3 As Single` (optional): If the EditingType of the new segment is msoEditingCorner, this argument specifies the horizontal distance, measured in points, from the upper-left corner of the document to the ending point of the new segment. If the EditingType of the new segment is msoEditingAuto, don't specify a value for this argument.
    - `Y3 As Single` (optional): If the EditingType of the new segment is msoEditingCorner, this argument specifies the vertical distance, measured in points, from the upper-left corner of the document to the ending point of the new segment. If the EditingType of the new segment is msoEditingAuto, don't specify a value for this argument.
