# GradientStop

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03BF-0000-0000-C000-000000000046}  

Represents one gradient stop.

**Remarks:** Gradients are a smooth transition from one color state to another. The endpoints of these sections are called stops.

**Example:**

```vba
Sub gradients()
 Set myDocument = ActivePresentation.Slides(1)
 Set GradientShapeFill = myDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 80).Fill
 With GradientShapeFill
 .ForeColor.RGB = RGB(0, 128, 128)
 .OneColorGradient msoGradientHorizontal, 1, 1
 .GradientStops.Insert RGB(255, 0, 0), 0.25
 .GradientStops.Insert RGB(0, 255, 0), 0.5
 .GradientStops.Insert RGB(0, 0, 255), 0.75
 End With
 GradientShapeFill.GradientStops.Delete (1)
End Sub
```

## Properties (5)

- `Application As Object  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Office application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the GradientStop object was created. Read-only.
- `Color As ColorFormat  (read-only)`  
  Gets a value representing the color of the gradient stop. Read-only.
- `Position As Single  (read/write)`  
  Gets or sets a value representing the position of a stop within the gradient expressed as a percent. Read/write.
- `Transparency As Single  (read/write)`  
  Gets or sets a value representing the opacity of the gradient fill expressed as a percent. Read/write.
