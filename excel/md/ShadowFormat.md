# ShadowFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000C031B-0000-0000-C000-000000000046}  

Represents shadow formatting for a shape.

**Remarks:** Use the Shadow property of the Shape object to return a ShadowFormat object.

**Example:**

```vba
Set myDocument = Worksheets(1)
With myDocument.Shapes.AddShape(msoShapeRectangle, _
 50, 50, 100, 200).Shadow
 .ForeColor.RGB = RGB(0, 0, 128)
 .OffsetX = 5
 .OffsetY = -3
 .Transparency = 0.5
 .Visible = True
End With
```

## Properties (14)

- `Application As Object  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `ForeColor As ColorFormat  (read/write)`  
  Returns or sets a ColorFormat object that represents the specified foreground fill or solid color.
- `Obscured As MsoTriState  (read/write)`  
  True if the shadow of the specified shape appears filled in and is obscured by the shape, even if the shape has no fill.
- `OffsetX As Single  (read/write)`  
  Returns or sets the horizontal offset of the shadow from the specified shape, in points. A positive value offsets the shadow to the right of the shape; a negative value offsets it to the left. Read/write Single.
- `OffsetY As Single  (read/write)`  
  Returns or sets the vertical offset of the shadow from the specified shape, in points. A positive value offsets the shadow down; a negative value offsets it up. Read/write Single.
- `Transparency As Single  (read/write)`  
  Returns or sets the degree of transparency of the specified fill as a value from 0.0 (opaque) through 1.0 (clear). Read/write Double.
- `Type As MsoShadowType  (read/write)`  
  Returns or sets an MsoShadowType value that represents the shadow format type.
- `Visible As MsoTriState  (read/write)`  
  Returns or sets an MsoTriState value that determines whether the object is visible. Read/write.
- `Style As MsoShadowStyle  (read/write)`  
  Returns or sets the style of the specified shadow. Read/write MsoShadowStyle.
- `Blur As Single  (read/write)`  
  Returns or sets the degree of blurriness of the specified shadow. Read/write Single.
- `Size As Single  (read/write)`  
  Returns or sets the size of the specified shadow. Read/write Single.
- `RotateWithShape As MsoTriState  (read/write)`  
  Returns or sets an MsoTriState value that represents whether to rotate the shadow when rotating the shape. Read/write.

## Methods (2)

- `IncrementOffsetX(Increment As Single)`  
  Changes the horizontal offset of the shadow by the specified number of points. Use the OffsetX property to set the absolute horizontal shadow offset.
    - `Increment As Single` (required): Specifies how far the shadow offset is to be moved horizontally, in points. A positive value moves the shadow to the right; a negative value moves it to the left.
- `IncrementOffsetY(Increment As Single)`  
  Changes the vertical offset of the shadow by the specified number of points. Use the OffsetY property to set the absolute vertical shadow offset.
    - `Increment As Single` (required): Specifies how far the shadow offset is to be moved vertically, in points. A positive value moves the shadow down; a negative value moves it up.
