# ChartFillFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {F152D349-7D20-4C01-A42B-2D6DE4F3891C}  

## Properties (16)

- `BackColor As ChartColorFormat  (read-only)`
- `ForeColor As ChartColorFormat  (read-only)`
- `GradientColorType As MsoGradientColorType  (read-only)`
- `GradientDegree As Single  (read-only)`
- `GradientStyle As MsoGradientStyle  (read-only)`
- `GradientVariant As Long  (read-only)`
- `Pattern As MsoPatternType  (read-only)`
- `PresetGradientType As MsoPresetGradientType  (read-only)`
- `PresetTexture As MsoPresetTexture  (read-only)`
- `TextureName As String  (read-only)`
- `TextureType As MsoTextureType  (read-only)`
- `Type As MsoFillType  (read-only)`
- `Visible As MsoTriState  (read/write)`
- `Application As Object  (read-only)`
- `Creator As Long  (read-only)`
- `Parent As Object  (read-only)`

## Methods (8)

- `OneColorGradient(Style As MsoGradientStyle, Variant As Long, Degree As Single)`
- `TwoColorGradient(Style As MsoGradientStyle, Variant As Long)`
- `PresetTextured(PresetTexture As MsoPresetTexture)`
- `Solid()`
- `Patterned(Pattern As MsoPatternType)`
- `UserPicture([PictureFile As Variant], [PictureFormat As Variant], [PictureStackUnit As Variant], [PicturePlacement As Variant])`
- `UserTextured(TextureFile As String)`
- `PresetGradient(Style As MsoGradientStyle, Variant As Long, PresetGradientType As MsoPresetGradientType)`
