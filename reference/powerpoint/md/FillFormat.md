# FillFormat

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {9149347E-5A91-11CF-8700-00AA0060263B}  

Represents fill formatting for a shape. A shape can have a solid, gradient, texture, pattern, picture, or semi-transparent fill.

**Remarks:** Many of the properties of the FillFormat object are read-only. To set one of these properties, you have to apply the corresponding method.

**Example:**

```vba
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes _

        .AddShape(msoShapeRectangle, 90, 90, 90, 80).Fill

    .ForeColor.RGB = RGB(0, 128, 128)

    .OneColorGradient msoGradientHorizontal, 1, 1

End With
```

## Properties (27)

- `Application As Object  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Creator As Long  (read-only)`  
  Returns a Long that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in Microsoft PowerPoint, this property returns the hexadecimal number 50575054. Read-only.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `BackColor As ColorFormat  (read/write)`  
  Returns or sets a ColorFormat object that represents the background color for the specified fill or patterned line. Read/write.
- `ForeColor As ColorFormat  (read/write)`  
  Returns or sets a ColorFormat object that represents the foreground color for the fill, line, or shadow. Read/write.
- `GradientColorType As MsoGradientColorType  (read-only)`  
  Returns the gradient color type for the specified fill. Read-only.
- `GradientDegree As Single  (read-only)`  
  Returns a value that indicates how dark or light a one-color gradient fill is. Read-only.
- `GradientStyle As MsoGradientStyle  (read-only)`  
  Returns the gradient style for the specified fill. Read-only.
- `GradientVariant As Long  (read-only)`  
  Returns the gradient variant for the specified fill as an integer value from 1 to 4 for most gradient fills. Read-only.
- `Pattern As MsoPatternType  (read-only)`  
  Sets or returns a value that represents the pattern applied to the specified fill. Read-only.
- `PresetGradientType As MsoPresetGradientType  (read-only)`  
  Returns the preset gradient type for the specified fill. Read-only.
- `PresetTexture As MsoPresetTexture  (read-only)`  
  Returns the preset texture for the specified fill. Read-only.
- `TextureName As String  (read-only)`  
  Returns the name of the custom texture file for the specified fill. Read-only.
- `TextureType As MsoTextureType  (read-only)`  
  Returns the texture type for the specified fill. Read-only.
- `Transparency As Single  (read/write)`  
  Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0 (opaque) and 1.0 (clear). Read/write.
- `Type As MsoFillType  (read-only)`  
  Represent the type of fill. Read-only.
- `Visible As MsoTriState  (read/write)`  
  Returns or sets the visibility of the specified object or the formatting applied to the specified object. Read/write.
- `GradientStops As GradientStops  (read-only)`  
  Returns the GradientStops collection associated with the specified fill format. Read-only.
- `TextureOffsetX As Single  (read/write)`  
  Returns or sets the horizontal offset of the texture from the origin in points. Read/write.
- `TextureOffsetY As Single  (read/write)`  
  Returns or sets the vertical offset of the texture from the origin in points. Read/write.
- `TextureAlignment As MsoTextureAlignment  (read/write)`  
  Returns or sets the alignment (the origin of the coordinate grid) for the tiling of the texture fill. Read/write.
- `TextureHorizontalScale As Single  (read/write)`  
  Returns or sets the horizontal scaling factor for the texture fill. Read/write.
- `TextureVerticalScale As Single  (read/write)`  
  Returns or sets the vertical scaling factor for the texture fill. Read/write.
- `TextureTile As MsoTriState  (read/write)`  
  Returns or sets whether the texture fill is tiled or centered. Read/write.
- `RotateWithObject As MsoTriState  (read/write)`  
  Returns or sets whether the fill rotates with the specified shape. Read/write.
- `PictureEffects As PictureEffects  (read-only)`  
  Returns an object that represents the picture or texture fill for the specified fill format. Read-only.
- `GradientAngle As Single  (read/write)`  
  Returns or sets the angle of the gradient fill for the specified fill format. Read/write.

## Methods (9)

- `Background()`  
  Specifies that the shape's fill should match the slide background. If you change the slide background after applying this method to a fill, the fill will also change.
- `OneColorGradient(Style As MsoGradientStyle, Variant As Long, Degree As Single)`  
  Sets the specified fill to a one-color gradient.
    - `Style As MsoGradientStyle` (required): The gradient style.
    - `Variant As Long` (required): The gradient variant. Can be a value from 1 to 4, corresponding to the four variants on the Gradient tab in the Shape Fill tab. If Style is msoGradientFromTitle or msoGradientFromCenter, this argument can be either 1 or 2.
    - `Degree As Single` (required): The gradient degree. Can be a value from 0.0 (dark) to 1.0 (light).
- `Patterned(Pattern As MsoPatternType)`  
  Sets the specified fill to a pattern.
    - `Pattern As MsoPatternType` (required): The pattern to be used for the specified fill. See Remarks for possible values.
- `PresetGradient(Style As MsoGradientStyle, Variant As Long, PresetGradientType As MsoPresetGradientType)`  
  Sets the specified fill to a preset gradient.
    - `Style As MsoGradientStyle` (required): The gradient style.
    - `Variant As Long` (required): The gradient variant. Can be a value from 1 to 4, corresponding to the four variants on the Gradient subtab on the Shape Fill tab. If Style is msoGradientFromTitle or msoGradientFromCenter, this argument can be either 1 or 2.
    - `PresetGradientType As MsoPresetGradientType` (required): The gradient type.
- `PresetTextured(PresetTexture As MsoPresetTexture)`  
  Sets the specified fill to a preset texture.
    - `PresetTexture As MsoPresetTexture` (required): The preset texture.
- `Solid()`  
  Sets the specified fill to a uniform color. Use this method to convert a gradient, textured, patterned, or background fill back to a solid fill.
- `TwoColorGradient(Style As MsoGradientStyle, Variant As Long)`  
  Sets the specified fill to a two-color gradient.
    - `Style As MsoGradientStyle` (required): The gradient style.
    - `Variant As Long` (required): The gradient variant. Can be from 1 to 4, corresponding to the four variants on the Gradient sub-tab on the Shape Fill tab. If Style is msoGradientFromTitle or msoGradientFromCenter, this argument can be either 1 or 2.
- `UserPicture(PictureFile As String)`  
  Fills the specified shape with one large image.
    - `PictureFile As String` (required): The name of the picture file.
- `UserTextured(TextureFile As String)`  
  Fills the specified shape with small tiles of an image.
