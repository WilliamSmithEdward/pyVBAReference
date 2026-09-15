# ShapeNode

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493487-5A91-11CF-8700-00AA0060263B}  

Represents the geometry and the geometry-editing properties of the nodes in a user-defined freeform.

**Remarks:** Nodes include the vertices between the segments of the freeform and the control points for curved segments. The ShapeNode object is a member of the ShapeNodes collection. The ShapeNodes collection contains all the nodes in a freeform.

**Example:**

```vba
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3)

    If .Nodes(1).EditingType = msoEditingCorner Then

        .Nodes.SetEditingType 1, msoEditingSmooth

    End If

End With
```

## Properties (6)

- `Application As Object  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Creator As Long  (read-only)`  
  Returns a Long that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in Microsoft PowerPoint, this property returns the hexadecimal number 50575054. Read-only.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `EditingType As MsoEditingType  (read-only)`  
  If the specified node is a vertex, this property returns a value that indicates how changes made to the node affect the two segments connected to the node. If the node is a control point for a curved segment, this property returns the editing type of the adjacent vertex. Read-only.
- `Points As Variant  (read-only)`  
  Returns a Variant that represents the position of the specified node as a coordinate pair. Read-only.
- `SegmentType As MsoSegmentType  (read-only)`  
  Returns a value that indicates whether the segment associated with the specified node is straight or curved. Read-only.
