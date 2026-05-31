# FillFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000C0314-0000-0000-C000-000000000046}  

Represents fill formatting for a shape.

**Remarks:** A shape can have a solid, gradient, texture, pattern, picture, or semi-transparent fill. Many of the properties of the FillFormat object are read-only. To set one of these properties, you have to apply the corresponding method.

**Example:**

```vba
Set myDocument = Worksheets(1)
With myDocument.Shapes.AddShape(msoShapeRectangle, _
 90, 90, 90, 80).Fill
 .ForeColor.RGB = RGB(0, 128, 128)
 .OneColorGradient msoGradientHorizontal, 1, 1
End With
```

## Properties (27)

- `Application As Object  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `BackColor As ColorFormat  (read/write)`  
  Returns or sets a ColorFormat object that represents the specified fill background color.
- `ForeColor As ColorFormat  (read/write)`  
  Returns or sets a ColorFormat object that represents the specified foreground fill or solid color.
- `GradientColorType As MsoGradientColorType  (read-only)`  
  Returns the gradient color type for the specified fill. Read-only MsoGradientColorType.
- `GradientDegree As Single  (read-only)`  
  Returns the gradient degree of the specified one-color shaded fill as a floating-point value from 0.0 (dark) through 1.0 (light). Read-only Single.
- `GradientStyle As MsoGradientStyle  (read-only)`  
  Returns the gradient style for the specified fill. Read-only MsoGradientStyle.
- `GradientVariant As Long  (read-only)`  
  Returns the shade variant for the specified fill as an integer value from 1 through 4. The values for this property correspond to the gradient variants (numbered from left to right and from top to bottom) on the Gradient tab in the Fill Effects dialog box. Read-only Long.
- `Pattern As MsoPatternType  (read-only)`  
  Returns or sets an MsoPatternType value that represents the fill pattern.
- `PresetGradientType As MsoPresetGradientType  (read-only)`  
  Returns the preset gradient type for the specified fill. Read-only MsoPresetGradientType.
- `PresetTexture As MsoPresetTexture  (read-only)`  
  Returns the preset texture for the specified fill. Read-only MsoPresetTexture.
- `TextureName As String  (read-only)`  
  Returns the name of the custom texture file for the specified fill. Read-only String.
- `TextureType As MsoTextureType  (read-only)`  
  Returns the texture type for the specified fill. Read-only MsoTextureType.
- `Transparency As Single  (read/write)`  
  Returns or sets the degree of transparency of the specified fill as a value from 0.0 (opaque) through 1.0 (clear). Read/write Double.
- `Type As MsoFillType  (read-only)`  
  Returns an MsoFillType value that represents the fill type.
- `Visible As MsoTriState  (read/write)`  
  Returns or sets an MsoTriState value that determines whether the object is visible. Read/write.
- `GradientStops As GradientStops  (read-only)`  
  Returns the end point for the gradient fill. Read-only.
- `TextureOffsetX As Single  (read/write)`  
  Returns the offset X value for the specified fill. Read/write Single.
- `TextureOffsetY As Single  (read/write)`  
  Returns the offset Y value for the specified fill. Read/write Single.
- `TextureAlignment As MsoTextureAlignment  (read/write)`  
  Returns or sets the alignment (the origin of the coordinate grid) for the tiling of the texture fill. Read/write.
- `TextureHorizontalScale As Single  (read/write)`  
  Returns or sets the horizontal scaling factor for the texture fill. Read/write Single.
- `TextureVerticalScale As Single  (read/write)`  
  Returns the texture vertical scale for the specified fill. Read/write Single.
- `TextureTile As MsoTriState  (read/write)`  
  Returns the texture tile style for the specified fill. Read/write MsoTriState.
- `RotateWithObject As MsoTriState  (read/write)`  
  Returns or sets if the fill style should rotate with the object. Read/write MsoTriState.
- `PictureEffects As PictureEffects  (read-only)`  
  Returns a PictureEffects object that represents the picture or texture fill for the specified fill format. Read-only.
- `GradientAngle As Single  (read/write)`  
  Returns or sets the angle of the gradient fill for the specified fill format. Read/write.

## Methods (8)

- `OneColorGradient(Style As MsoGradientStyle, Variant As Long, Degree As Single)`  
  Sets the specified fill to a one-color gradient.
    - `Style As MsoGradientStyle` (required): The gradient style.
    - `Variant As Long` (required): The gradient variant. Can be a value from 1 through 4, corresponding to one of the four variants on the Gradient tab in the Fill Effects dialog box. If _GradientStyle_ is msoGradientFromCenter, the _Variant_ argument can only be 1 or 2.
    - `Degree As Single` (required): The gradient degree. Can be a value from 0.0 (dark) through 1.0 (light).
- `Patterned(Pattern As MsoPatternType)`  
  Sets the specified fill to a pattern.
    - `Pattern As MsoPatternType` (required): The type of pattern.
- `PresetGradient(Style As MsoGradientStyle, Variant As Long, PresetGradientType As MsoPresetGradientType)`  
  Sets the specified fill to a preset gradient.
    - `Style As MsoGradientStyle` (required): The gradient style.
    - `Variant As Long` (required): The gradient variant. Can be a value from 1 through 4, corresponding to one of the four variants on the Gradient tab in the Fill Effects dialog box. If _Style_ is msoGradientFromCenter, the _Variant_ argument can only be 1 or 2.
    - `PresetGradientType As MsoPresetGradientType` (required): The preset gradient type.
- `PresetTextured(PresetTexture As MsoPresetTexture)`  
  Sets the specified fill format to a preset texture.
    - `PresetTexture As MsoPresetTexture` (required): The type of texture to apply.
- `Solid()`  
  Sets the specified fill to a uniform color. Use this method to convert a gradient, textured, patterned, or background fill back to a solid fill.
- `TwoColorGradient(Style As MsoGradientStyle, Variant As Long)`  
  Sets the specified fill to a two-color gradient.
    - `Style As MsoGradientStyle` (required): The gradient style.
    - `Variant As Long` (required): The gradient variant. Can be a value from 1 through 4, corresponding to one of the four variants on the Gradient tab in the Fill Effects dialog box. If _Style_ is msoGradientFromCenter, the _Variant_ argument can only be 1 or 2.
- `UserPicture(PictureFile As String)`  
  Fills the specified shape with an image.
    - `PictureFile As String` (required): The file path to the picture file, or the name of the picture file if a link to the picture file currently exists.
- `UserTextured(TextureFile As String)`  
  Fills the specified shape with small tiles of an image. If you want to fill the shape with one large image, use the UserPicture method.
    - `TextureFile As String` (required): The name of the picture file.
