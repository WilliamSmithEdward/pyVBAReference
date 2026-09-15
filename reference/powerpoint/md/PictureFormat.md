# PictureFormat

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {9149347D-5A91-11CF-8700-00AA0060263B}  

Contains properties and methods that apply to pictures and OLE objects.

**Example:**

```vba
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(1).PictureFormat

    .Brightness = 0.3

    .Contrast = 0.7

    .ColorType = msoPictureGrayScale

    .CropBottom = 18

End With
```

## Properties (13)

- `Application As Object  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Creator As Long  (read-only)`  
  Returns a Long that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in Microsoft PowerPoint, this property returns the hexadecimal number 50575054. Read-only.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Brightness As Single  (read/write)`  
  Returns or sets the brightness of the specified picture or OLE object. Read/write.
- `ColorType As MsoPictureColorType  (read/write)`  
  Returns or sets the type of color transformation applied to the specified picture or OLE object. Read/write.
- `Contrast As Single  (read/write)`  
  Returns or sets the contrast for the specified picture or OLE object.
- `CropBottom As Single  (read/write)`  
  Returns or sets the number of points that are cropped off the bottom of the specified picture or OLE object. Read/write.
- `CropLeft As Single  (read/write)`  
  Returns or sets the number of points that are cropped off the left side of the specified picture or OLE object. Read/write.
- `CropRight As Single  (read/write)`  
  Returns or sets the number of points that are cropped off the right side of the specified picture or OLE object. Read/write.
- `CropTop As Single  (read/write)`  
  Returns or sets the number of points that are cropped off the top of the specified picture or OLE object. Read/write.
- `TransparencyColor As MsoRGBType  (read/write)`  
  Returns or sets the transparent color for the specified picture as a red-green-blue (RGB) value. Read/write.
- `TransparentBackground As MsoTriState  (read/write)`  
  Determines whether parts of the picture that are the color defined as the transparent color appear transparent. Applies to bitmaps only. Read/write.
- `Crop As Crop  (read-only)`  
  Returns or sets the number of points that are cropped off the specified picture or OLE object. Read-only.

## Methods (2)

- `IncrementBrightness(Increment As Single)`  
  Changes the brightness of the picture by the specified amount.
    - `Increment As Single` (required): Specifies how much to change the value of the Brightness property for the picture. A positive value makes the picture brighter; a negative value makes the picture darker.
- `IncrementContrast(Increment As Single)`  
  Changes the contrast of the picture by the specified amount.
    - `Increment As Single` (required): Specifies how much to change the value of the Contrast property for the picture. A positive value increases the contrast; a negative value decreases the contrast.
