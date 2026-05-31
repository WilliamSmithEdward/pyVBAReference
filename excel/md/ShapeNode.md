# ShapeNode

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000C0318-0000-0000-C000-000000000046}  

Represents the geometry and the geometry-editing properties of the nodes in a user-defined freeform.

**Remarks:** Nodes include the vertices between the segments of the freeform and the control points for curved segments. The ShapeNode object is a member of the ShapeNodes collection. The ShapeNodes collection contains all the nodes in a freeform.

**Example:**

```vba
Set myDocument = Worksheets(1)
With myDocument.Shapes(3)
 If .Nodes(1).EditingType = msoEditingCorner Then
 .Nodes.SetEditingType 1, msoEditingSmooth
 End If
End With
```

## Properties (6)

- `Application As Object  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `EditingType As MsoEditingType  (read-only)`  
  If the specified node is a vertex, this property returns a value that indicates how changes made to the node affect the two segments connected to the node. Read-only MsoEditingType.
- `Points As Variant  (read-only)`  
  Returns the position of the specified node as a coordinate pair. Each coordinate is expressed in points. Read-only Variant.
- `SegmentType As MsoSegmentType  (read-only)`  
  Returns a value that indicates whether the segment associated with the specified node is straight or curved. If the specified node is a control point for a curved segment, this property returns msoSegmentCurve. Read-only MsoSegmentType.
