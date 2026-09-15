# ShapeNode

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209CD-0000-0000-C000-000000000046}  

Represents the geometry and the geometry-editing properties of the nodes in a user-defined freeform. Nodes include the vertices between the segments of the freeform and the control points for curved segments. The ShapeNode object is a member of the ShapeNodes collection. The ShapeNodes collection contains all the nodes in a freeform.

**Remarks:** Use Nodes (Index), where Index is the node index number, to return a single ShapeNode object. If node one in shape three on the active document is a corner point, the following example makes it a smooth point. For this example to work, shape three must be a freeform.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ShapeNode object.
- `EditingType As MsoEditingType  (read-only)`  
  If the specified node is a vertex, this property returns a value that indicates how changes made to the node affect the two segments connected to the node. Read-only MsoEditingType. .
- `Points As Variant  (read-only)`  
  Returns the position of the specified node as a coordinate pair. Read-only Variant.
- `SegmentType As MsoSegmentType  (read-only)`  
  Returns a value that indicates whether the segment associated with the specified node is straight or curved. Read-only MsoSegmentType.
