# IChartFillFormat

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024435-0001-0000-C000-000000000046}  

## Properties (16)

- `Application As HRESULT  (read-only)`
- `Creator As HRESULT  (read-only)`
- `Parent As HRESULT  (read-only)`
- `BackColor As HRESULT  (read-only)`
- `ForeColor As HRESULT  (read-only)`
- `GradientColorType As HRESULT  (read-only)`
- `GradientDegree As HRESULT  (read-only)`
- `GradientStyle As HRESULT  (read-only)`
- `GradientVariant As HRESULT  (read-only)`
- `Pattern As HRESULT  (read-only)`
- `PresetGradientType As HRESULT  (read-only)`
- `PresetTexture As HRESULT  (read-only)`
- `TextureName As HRESULT  (read-only)`
- `TextureType As HRESULT  (read-only)`
- `Type As HRESULT  (read-only)`
- `Visible As HRESULT  (read/write)`

## Methods (8)

- `OneColorGradient(Style As MsoGradientStyle, Variant As Long, Degree As Single)`
- `TwoColorGradient(Style As MsoGradientStyle, Variant As Long)`
- `PresetTextured(PresetTexture As MsoPresetTexture)`
- `Solid()`
- `Patterned(Pattern As MsoPatternType)`
- `UserPicture([PictureFile As Variant], [PictureFormat As Variant], [PictureStackUnit As Variant], [PicturePlacement As Variant])`
- `UserTextured(TextureFile As String)`
- `PresetGradient(Style As MsoGradientStyle, Variant As Long, PresetGradientType As MsoPresetGradientType)`
