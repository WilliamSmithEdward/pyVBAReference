# GradientStops

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03C0-0000-0000-C000-000000000046}  

Contains a collection of GradientStop objects.

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
End Sub
```

## Properties (5)

- `Application As Object  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Office application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the GradientStops object was created. Read-only.
- `Item As GradientStop  (read-only)`  
  Gets a GradientStop object from a GradientStops collection. Read-only.
- `Count As Long  (read-only)`  
  Gets a Long indicating the number of items in the GradientStops collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`

## Methods (3)

- `Delete([Index As Long])`  
  Removes a gradient stop.
    - `Index As Long` (optional): The index number of the gradient stop.
- `Insert(RGB As MsoRGBType, Position As Single, [Transparency As Single], [Index As Long])`  
  Adds a stop to a gradient.
    - `RGB As MsoRGBType` (required): Specifies the color at the gradient stop.
    - `Position As Single` (required): Specifies the position of the stop within the gradient expressed as a percent.
    - `Transparency As Single` (optional): Specifies the opacity of color at the gradient stop.
    - `Index As Long` (optional): The index number of the stop.
- `Insert2(RGB As MsoRGBType, Position As Single, [Transparency As Single], [Index As Long], [Brightness As Single])`  
  Adds a stop to a gradient, and specifies the brightness, as well as the transparency, of the color.
    - `RGB As MsoRGBType` (required): Specifies the color at the gradient stop.
    - `Position As Single` (required): Specifies the position of the stop within the gradient expressed as a percent.
    - `Transparency As Single` (optional): Specifies the opacity of the color at the gradient stop.
    - `Index As Long` (optional): The index number of the gradient stop.
    - `Brightness As Single` (optional): Specifies the brightness of the color at the gradient stop.
