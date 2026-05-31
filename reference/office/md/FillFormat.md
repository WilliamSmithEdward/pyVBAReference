# FillFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0314-0000-0000-C000-000000000046}  

## Properties (27)

- `Application As Object  (read-only)`
- `Creator As Long  (read-only)`
- `Parent As Object  (read-only)`
- `BackColor As ColorFormat  (read/write)`
- `ForeColor As ColorFormat  (read/write)`
- `GradientColorType As MsoGradientColorType  (read-only)`
- `GradientDegree As Single  (read-only)`
- `GradientStyle As MsoGradientStyle  (read-only)`
- `GradientVariant As Long  (read-only)`
- `Pattern As MsoPatternType  (read-only)`
- `PresetGradientType As MsoPresetGradientType  (read-only)`
- `PresetTexture As MsoPresetTexture  (read-only)`
- `TextureName As String  (read-only)`
- `TextureType As MsoTextureType  (read-only)`
- `Transparency As Single  (read/write)`
- `Type As MsoFillType  (read-only)`
- `Visible As MsoTriState  (read/write)`
- `GradientStops As GradientStops  (read-only)`
- `TextureOffsetX As Single  (read/write)`
- `TextureOffsetY As Single  (read/write)`
- `TextureAlignment As MsoTextureAlignment  (read/write)`
- `TextureHorizontalScale As Single  (read/write)`
- `TextureVerticalScale As Single  (read/write)`
- `TextureTile As MsoTriState  (read/write)`
- `RotateWithObject As MsoTriState  (read/write)`
- `PictureEffects As PictureEffects  (read-only)`
- `GradientAngle As Single  (read/write)`

## Methods (9)

- `Background()`
- `OneColorGradient(Style As MsoGradientStyle, Variant As Long, Degree As Single)`
- `Patterned(Pattern As MsoPatternType)`
- `PresetGradient(Style As MsoGradientStyle, Variant As Long, PresetGradientType As MsoPresetGradientType)`
- `PresetTextured(PresetTexture As MsoPresetTexture)`
- `Solid()`
- `TwoColorGradient(Style As MsoGradientStyle, Variant As Long)`
- `UserPicture(PictureFile As String)`
- `UserTextured(TextureFile As String)`
