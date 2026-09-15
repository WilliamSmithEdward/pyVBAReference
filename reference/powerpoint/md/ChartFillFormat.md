# ChartFillFormat

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {92D41A5B-F07E-4CA4-AF6F-BEF486AA4E6F}  

## Properties (16)

- `BackColor As ChartColorFormat  (read-only)`
- `ForeColor As ChartColorFormat  (read-only)`
- `GradientDegree As Single  (read-only)`
- `TextureName As String  (read-only)`
- `Creator As Long  (read-only)`
- `Parent As Object  (read-only)`
- `Application As Application  (read-only)`
- `GradientColorType As MsoGradientColorType  (read-only)`
- `GradientStyle As MsoGradientStyle  (read-only)`
- `GradientVariant As Long  (read-only)`
- `Pattern As MsoPatternType  (read-only)`
- `PresetGradientType As MsoPresetGradientType  (read-only)`
- `PresetTexture As MsoPresetTexture  (read-only)`
- `TextureType As MsoTextureType  (read-only)`
- `Type As MsoFillType  (read-only)`
- `Visible As MsoTriState  (read/write)`

## Methods (8)

- `Solid()`
- `UserTextured(TextureFile As String)`
- `OneColorGradient(Style As MsoGradientStyle, Variant As Long, Degree As Single)`
- `Patterned(Pattern As MsoPatternType)`
- `PresetGradient(Style As MsoGradientStyle, Variant As Long, PresetGradientType As MsoPresetGradientType)`
- `PresetTextured(PresetTexture As MsoPresetTexture)`
- `TwoColorGradient(Style As MsoGradientStyle, Variant As Long)`
- `UserPicture([PictureFile As Variant], [PictureFormat As Variant], [PictureStackUnit As Variant], [PicturePlacement As Variant])`
