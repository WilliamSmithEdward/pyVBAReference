# FillFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209C8-0000-0000-C000-000000000046}  

Represents fill formatting for a shape. A shape can have a solid, gradient, texture, pattern, picture, or semi-transparent fill.

**Remarks:** Use the Fill property to return a FillFormat object. The following example adds a rectangle to the active document and then sets the gradient and color for the rectangle's fill. Many of the properties of the FillFormat object are read-only. To set one of these properties, you have to apply the corresponding method.

## Properties (27)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified FillFormat object.
- `BackColor As ColorFormat  (read-only)`  
  Returns or sets a ColorFormat object that represents the background color for the fill Read/write.
- `ForeColor As ColorFormat  (read-only)`  
  Returns or sets a ColorFormat object that represents the foreground color for the fill. Read/write.
- `GradientColorType As MsoGradientColorType  (read-only)`  
  Returns the gradient color type for the specified fill. Read-only MsoGradientColorType.
- `GradientDegree As Single  (read-only)`  
  Returns a value that indicates how dark or light a one-color gradient fill is. Read-only Single.
- `GradientStyle As MsoGradientStyle  (read-only)`  
  Returns the gradient style for the specified fill. Read-only MsoGradientStyle.
- `GradientVariant As Long  (read-only)`  
  Returns the gradient variant for the specified fill as an integer value from 1 to 4 for most gradient fills. Read-only Long.
- `Pattern As MsoPatternType  (read-only)`  
  Returns or sets a MsoPatternType constant that represents the pattern applied to the specified fill or line. Read-only.
- `PresetGradientType As MsoPresetGradientType  (read-only)`  
  Returns the preset gradient type for the specified fill. Read-only MsoPresetGradientType.
- `PresetTexture As MsoPresetTexture  (read-only)`  
  Returns the preset texture for the specified fill. Read-only MsoPresetTexture.
- `TextureName As String  (read-only)`  
  Returns the name of the custom texture file for the specified fill. Read-only String.
- `TextureType As MsoTextureType  (read-only)`  
  Returns the texture type for the specified fill. Read-only MsoTextureType.
- `Transparency As Single  (read/write)`  
  Returns or sets the degree of transparency of the specified fillfor a shape as a value between 0.0 (opaque) and 1.0 (clear). Read/write Single.
- `Type As MsoFillType  (read-only)`  
  Returns the shape fill format type. Read-only MsoFillType.
- `Visible As MsoTriState  (read/write)`  
  True if the specified object, or the formatting applied to it, is visible. Read/write MsoTriState.
- `GradientStops As GradientStops  (read-only)`  
  Returns the GradientStops collection associated with the specified fill format. Read-only.
- `TextureOffsetX As Single  (read/write)`  
  Returns or sets a Long that specifies the horizontal offset of the texture from the origin in points. Read/write.
- `TextureOffsetY As Single  (read/write)`  
  Returns or sets a Long that specifies the vertical offset of the texture from the origin in points. Read/write.
- `TextureAlignment As MsoTextureAlignment  (read/write)`  
  Returns or sets the alignment (the origin of the coordinate grid) for the tiling of the texture fill. Read/write.
- `TextureHorizontalScale As Single  (read/write)`  
  Returns or sets a Single that specifies the horizontal scaling factor for the texture fill. Read/write.
- `TextureVerticalScale As Single  (read/write)`  
  Returns or sets a Single that specifies the vertical scaling factor for the texture fill. Read/write.
- `TextureTile As MsoTriState  (read/write)`  
  Returns or sets whether the texture fill is tiled or centered. Read/write.
- `RotateWithObject As MsoTriState  (read/write)`  
  Returns or sets whether the fill rotates with the specified shape. Read/write.
- `PictureEffects As PictureEffects  (read-only)`  
  Returns a PictureEffects object that can be used to apply picture effects to the specified fill formatting. Read-only.
- `GradientAngle As Single  (read/write)`  
  Returns or sets the angle of the gradient fill for the specified fill format. Read/write.

## Methods (8)

- `OneColorGradient(Style As MsoGradientStyle, Variant As Long, Degree As Single)`  
  Sets the specified fill to a one-color gradient.
    - `Style As MsoGradientStyle` (required): The gradient style. Can be any MsoGradientStyle constant except msoGradientFromTitle which applies only to Microsoft PowerPoint.
    - `Variant As Long` (required): The gradient variant. Can be a value from 1 to 4, corresponding to the four variants on the Gradient tab in the Fill Effects dialog box. If Style is msoGradientFromCenter, this argument can be either 1 or 2.
    - `Degree As Single` (required): The gradient degree. Can be a value from 0.0 (dark) to 1.0 (light).
- `Patterned(Pattern As MsoPatternType)`  
  Sets the specified fill to a pattern.
    - `Pattern As MsoPatternType` (required): The pattern to be used for the specified fill.
- `PresetGradient(Style As MsoGradientStyle, Variant As Long, PresetGradientType As MsoPresetGradientType)`  
  Sets the specified fill to a preset gradient.
    - `Style As MsoGradientStyle` (required): The gradient style. Can be any MsoGradientStyle constant except msoGradientFromTitle, which applies only to Microsoft PowerPoint.
    - `Variant As Long` (required): The gradient variant. Can be a value from 1 to 4, corresponding to the four variants on the Gradient tab in the Fill Effects dialog box. If _Style_ is msoGradientFromCenter, this argument can be either 1 or 2.
    - `PresetGradientType As MsoPresetGradientType` (required): The gradient type.
- `PresetTextured(PresetTexture As MsoPresetTexture)`  
  Sets the specified fill to a preset texture.
    - `PresetTexture As MsoPresetTexture` (required): The preset texture.
- `Solid()`  
  Sets the specified fill to a uniform color.
- `TwoColorGradient(Style As MsoGradientStyle, Variant As Long)`  
  Sets the specified fill to a two-color gradient.
    - `Style As MsoGradientStyle` (required): The gradient style. Can be any MsoGradientStyle constant except msoGradientFromTitle which applies only to Microsoft PowerPoint.
    - `Variant As Long` (required): The gradient variant. Can be a value from 1 to 4, corresponding to the four variants on the Gradient tab in the Fill Effects dialog box. If _Style_ is msoGradientFromCenter, this argument can be either 1 or 2.
- `UserPicture(PictureFile As String)`  
  Fills the specified shape with one large image.
    - `PictureFile As String` (required): The name of the picture file.
- `UserTextured(TextureFile As String)`  
  Fills the specified shape with small tiles of an image.
