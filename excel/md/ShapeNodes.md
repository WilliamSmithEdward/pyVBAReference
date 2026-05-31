# ShapeNodes

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000C0319-0000-0000-C000-000000000046}  

A collection of all the ShapeNode objects in the specified freeform.

**Remarks:** Each ShapeNode object represents either a node between segments in a freeform or a control point for a curved segment of a freeform. You can create a freeform manually or by using the BuildFreeform and ConvertToShape methods.

**Example:**

```vba
Set myDocument = Worksheets(1)
myDocument.Shapes(3).Nodes.Delete 4
```

## Properties (5)

- `Application As Object  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns an Integer value that represents the number of objects in the collection.
- `_NewEnum As IUnknown  (read-only)`

## Methods (6)

- `Item(Index As Variant) As ShapeNode`  
  Returns a single object from a collection.
    - `Index As Variant` (required): The name or index number for the object.
- `Delete(Index As Long)`  
  Deletes the object.
    - `Index As Long` (required): The index of the object to delete.
- `Insert(Index As Long, SegmentType As MsoSegmentType, EditingType As MsoEditingType, X1 As Single, Y1 As Single, [X2 As Single], [Y2 As Single], [X3 As Single], [Y3 As Single])`  
  Inserts a node into a freeform shape.
    - `Index As Long` (required): Long. The number of the shape node after which to insert a new node.
    - `SegmentType As MsoSegmentType` (required): The segment type.
    - `EditingType As MsoEditingType` (required): The editing type.
    - `X1 As Single` (required): If the _EditingType_ of the new segment is msoEditingAuto, this argument specifies the horizontal distance, measured in points, from the upper-left corner of the document to the end point of the new segment. If the _EditingType_ of the new node is msoEditingCorner, this argument specifies the horizontal distance, measured in points, from the upper-left corner of the document to the first control point for the new segment.
    - `Y1 As Single` (required): If the _EditingType_ of the new segment is msoEditingAuto, this argument specifies the vertical distance, measured in points, from the upper-left corner of the document to the end point of the new segment. If the _EditingType_ of the new node is msoEditingCorner, this argument specifies the vertical distance, measured in points, from the upper-left corner of the document to the first control point for the new segment.
    - `X2 As Single` (optional): If the _EditingType_ of the new segment is msoEditingCorner, this argument specifies the horizontal distance, measured in points, from the upper-left corner of the document to the second control point for the new segment. If the _EditingType_ of the new segment is msoEditingAuto, don't specify a value for this argument.
    - `Y2 As Single` (optional): If the _EditingType_ of the new segment is msoEditingCorner, this argument specifies the vertical distance, measured in points, from the upper-left corner of the document to the second control point for the new segment. If the _EditingType_ of the new segment is msoEditingAuto, don't specify a value for this argument.
    - `X3 As Single` (optional): If the _EditingType_ of the new segment is msoEditingCorner, this argument specifies the horizontal distance, measured in points, from the upper-left corner of the document to the end point of the new segment. If the _EditingType_ of the new segment is msoEditingAuto, don't specify a value for this argument.
    - `Y3 As Single` (optional): If the _EditingType_ of the new segment is msoEditingCorner, this argument specifies the vertical distance, measured in points, from the upper-left corner of the document to the end point of the new segment. If the _EditingType_ of the new segment is msoEditingAuto, don't specify a value for this argument.
- `SetEditingType(Index As Long, EditingType As MsoEditingType)`  
  Sets the editing type of the node specified by _Index_. If the node is a control point for a curved segment, this method sets the editing type of the node adjacent to it that joins two segments. Note that, depending on the editing type, this method may affect the position of adjacent nodes.
    - `Index As Long` (required): The node whose editing type is to be set.
    - `EditingType As MsoEditingType` (required): The editing property of the vertex.
- `SetPosition(Index As Long, X1 As Single, Y1 As Single)`  
  Sets the location of the node specified by _Index_. Note that, depending on the editing type of the node, this method may affect the position of adjacent nodes.
    - `Index As Long` (required): The node whose position is to be set.
    - `X1 As Single` (required): The position (in points) of the new node relative to the upper-left corner of the document.
    - `Y1 As Single` (required): The position (in points) of the new node relative to the upper-left corner of the document.
- `SetSegmentType(Index As Long, SegmentType As MsoSegmentType)`  
  Sets the segment type of the segment that follows the node specified by _Index_. If the node is a control point for a curved segment, this method sets the segment type for that curve. Note that this may affect the total number of nodes by inserting or deleting adjacent nodes.
    - `Index As Long` (required): The node whose segment type is to be set.
    - `SegmentType As MsoSegmentType` (required): Specifies if the segment is straight or curved.
