# ShadowFormat

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493480-5A91-11CF-8700-00AA0060263B}  

Represents shadow formatting for a shape.

**Example:**

```vba
Set myDocument = ActivePresentation.Slides(1)

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
  Returns an Application object that represents the creator of the specified object.
- `Creator As Long  (read-only)`  
  Returns a Long that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in Microsoft PowerPoint, this property returns the hexadecimal number 50575054. Read-only.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `ForeColor As ColorFormat  (read/write)`  
  Returns or sets a ColorFormat object that represents the foreground color for the fill, line, or shadow. Read/write.
- `Obscured As MsoTriState  (read/write)`  
  Determines whether the shadow of the specified shape appears filled in and is obscured by the shape. Read/write.
- `OffsetX As Single  (read/write)`  
  Returns or sets the horizontal offset of the shadow from the specified shape, in points. Read/write.
- `OffsetY As Single  (read/write)`  
  Returns or sets the vertical offset of the shadow from the specified shape, in points. Read/write.
- `Transparency As Single  (read/write)`  
  Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0 (opaque) and 1.0 (clear). Read/write.
- `Type As MsoShadowType  (read/write)`  
  Represents the type of shadow. Read/write.
- `Visible As MsoTriState  (read/write)`  
  Returns or sets the visibility of the specified object or the formatting applied to the specified object. Read/write.
- `Style As MsoShadowStyle  (read/write)`  
  Returns or sets the shadow style. Read/write.
- `Blur As Single  (read/write)`  
  Returns or sets the blur radius of the specified shadow, in points. Read/write.
- `Size As Single  (read/write)`  
  Returns or sets the size of the specified shadow as a percentage of the shape size, from 0 to 200. Read/write.
- `RotateWithShape As MsoTriState  (read/write)`  
  Returns or sets whether the specified shadow rotates when the shape it is associated with rotates. Read/write.

## Methods (2)

- `IncrementOffsetX(Increment As Single)`  
  Changes the horizontal offset of the shadow by the specified number of points.
    - `Increment As Single` (required): Specifies how far the shadow offset is to be moved horizontally, in points. A positive value moves the shadow to the right; a negative value moves it to the left.
- `IncrementOffsetY(Increment As Single)`  
  Changes the vertical offset of the shadow by the specified number of points.
    - `Increment As Single` (required): Specifies how far the shadow offset is to be moved vertically, in points. A positive value moves the shadow down; a negative value moves it up.
