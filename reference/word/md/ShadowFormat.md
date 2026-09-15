# ShadowFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209CC-0000-0000-C000-000000000046}  

Represents shadow formatting for a shape.

**Remarks:** Use the Shadow property to return a ShadowFormat object. The following example adds a shadowed rectangle to the active document. The semitransparent, blue shadow is offset 5 points to the right of the rectangle and 3 points above it.

## Properties (14)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ShadowFormat object.
- `ForeColor As ColorFormat  (read-only)`  
  Returns or sets a ColorFormat object that represents the foreground color for the fill, line, or shadow. Read/write.
- `Obscured As MsoTriState  (read/write)`  
  MsoTrue if the shadow of the specified shape appears filled in and is obscured by the shape, even if the shape has no fill. MsoFalse if the shadow has no fill and the outline of the shadow is visible through the shape if the shape has no fill. Read/write MsoTriState.
- `OffsetX As Single  (read/write)`  
  Returns or sets the horizontal offset (in points) of the shadow from the specified shape. A positive value offsets the shadow to the right of the shape; a negative value offsets it to the left. Read/write Single.
- `OffsetY As Single  (read/write)`  
  Returns or sets the vertical offset (in points) of the shadow from the specified shape. Read/write Single.
- `Transparency As Single  (read/write)`  
  Returns or sets the degree of transparency of the specified shadow as a value between 0.0 (opaque) and 1.0 (clear). Read/write Single.
- `Type As MsoShadowType  (read/write)`  
  Returns or sets the shape shadow type. Read/write MsoShadowType.
- `Visible As MsoTriState  (read/write)`  
  True if the specified object, or the formatting applied to it, is visible. Read/write MsoTriState.
- `Style As MsoShadowStyle  (read/write)`  
  Returns or sets a MsoShadowType that represents the type of shadow formatting to apply to a shape. Read/write.
- `Blur As Single  (read/write)`  
  Returns or sets a Single that represents the blur level for a shadow format. Read/write.
- `Size As Single  (read/write)`  
  Returns or sets a Single that represents the width of the shadow. Read/write.
- `RotateWithShape As MsoTriState  (read/write)`  
  Returns or sets an MsoTriState that represents whether to rotate the shadow when rotating the shape. Read/write.

## Methods (2)

- `IncrementOffsetX(Increment As Single)`  
  Changes the horizontal offset of the shadow by the specified number of points.
    - `Increment As Single` (required): Specifies how far the shadow offset is to be moved horizontally, in points. A positive value moves the shadow to the right; a negative value moves it to the left.
- `IncrementOffsetY(Increment As Single)`  
  Changes the vertical offset of the shadow by the specified number of points.
    - `Increment As Single` (required): Specifies how far the shadow offset is to be moved vertically, in points. A positive value moves the shadow down; a negative value moves it up.
