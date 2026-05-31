# ChartFillFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C171C-0000-0000-C000-000000000046}  

## Properties (16)

- `BackColor As ChartColorFormat  (read-only)`
- `ForeColor As ChartColorFormat  (read-only)`
- `GradientColorType As Long  (read-only)`
- `GradientDegree As Single  (read-only)`
- `GradientStyle As Long  (read-only)`
- `GradientVariant As Long  (read-only)`
- `Pattern As Long  (read-only)`
- `PresetGradientType As Long  (read-only)`
- `PresetTexture As Long  (read-only)`
- `TextureName As String  (read-only)`
- `TextureType As Long  (read-only)`
- `Type As Long  (read-only)`
- `Visible As Long  (read/write)`
- `Application As Object  (read-only)`
- `Creator As Long  (read-only)`
- `Parent As Object  (read-only)`

## Methods (8)

- `OneColorGradient(Style As Long, Variant As Long, Degree As Single)`
- `TwoColorGradient(Style As Long, Variant As Long)`
- `PresetTextured(PresetTexture As Long)`
- `Solid()`
- `Patterned(Pattern As Long)`
- `UserPicture(PictureFile As Variant, PictureFormat As Variant, PictureStackUnit As Variant, PicturePlacement As Variant)`
- `UserTextured(TextureFile As String)`
- `PresetGradient(Style As Long, Variant As Long, PresetGradientType As Long)`
