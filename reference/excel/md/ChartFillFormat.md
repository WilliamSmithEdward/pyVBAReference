# ChartFillFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024435-0000-0000-C000-000000000046}  

Represents fill formatting.

**Remarks:** Use the Fill property to return the ChartFillFormat object.

**Example:**

```vba
With myChart.ChartArea.Fill
    .Visible = True
    .ForeColor.SchemeColor = 15
    .BackColor.SchemeColor = 17
    .TwoColorGradient msoGradientHorizontal, 1
End With
```

## Properties (16)

- `Application As Application  (read-only)`
- `Creator As XlCreator  (read-only)`
- `Parent As Object  (read-only)`
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

## Methods (8)

- `OneColorGradient(Style As MsoGradientStyle, Variant As Long, Degree As Single)`
- `TwoColorGradient(Style As MsoGradientStyle, Variant As Long)`
- `PresetTextured(PresetTexture As MsoPresetTexture)`
- `Solid()`
- `Patterned(Pattern As MsoPatternType)`
- `UserPicture([PictureFile As Variant], [PictureFormat As Variant], [PictureStackUnit As Variant], [PicturePlacement As Variant])`
- `UserTextured(TextureFile As String)`
- `PresetGradient(Style As MsoGradientStyle, Variant As Long, PresetGradientType As MsoPresetGradientType)`
