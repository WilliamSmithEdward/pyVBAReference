# SoftEdgeFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03BC-0000-0000-C000-000000000046}  

Represents the soft edge effect in Office graphics.

**Remarks:** The soft edge effect creates a mask around the edge of an object and blends the object with the transparent edge. The result is a faded or "feathered" edge.

**Example:**

```vba
With ActivePresentation.Slides(1).Shapes(2)
 With .Text.Font
 .Size = 32
 .Name = "Palatino"
 .Softedgeformat = msosoftedge6
 End With
End With
```

## Properties (4)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the SoftEdgeFormat object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the SoftEdgeFormat object was created. Read-only.
- `Type As MsoSoftEdgeType  (read/write)`  
  Gets or sets the type of the SoftEdgeFormat object. Read/write.
- `Radius As Single  (read/write)`  
  Gets or sets the size, measured in points, of the soft edge effect of the shape. Read/write.
