# PictureFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000C031A-0000-0000-C000-000000000046}  

Contains properties and methods that apply to pictures and OLE objects.

**Remarks:** The LinkFormat object contains properties and methods that apply to linked OLE objects only. The OLEFormat object contains properties and methods that apply to OLE objects whether or not they're linked.

**Example:**

```vba
Set myDocument = Worksheets(1)
With myDocument.Shapes(1).PictureFormat
 .Brightness = 0.3
 .Contrast = 0.7
 .ColorType = msoPictureGrayScale
 .CropBottom = 18
```

## Properties (13)

- `Application As Object  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Brightness As Single  (read/write)`  
  Returns or sets the brightness of the specified picture or OLE object. The value for this property must be a number from 0.0 (dimmest) to 1.0 (brightest). Read/write Single.
- `ColorType As MsoPictureColorType  (read/write)`  
  Returns or sets the type of color transformation applied to the specified picture or OLE object. Read/write.
- `Contrast As Single  (read/write)`  
  Returns or sets the contrast for the specified picture or OLE object. The value for this property must be a number from 0.0 (the least contrast) to 1.0 (the greatest contrast). Read/write Single.
- `CropBottom As Single  (read/write)`  
  Returns or sets the number of points that are cropped off the bottom of the specified picture or OLE object. Read/write Single.
- `CropLeft As Single  (read/write)`  
  Returns or sets the number of points that are cropped off the left side of the specified picture or OLE object. Read/write Single.
- `CropRight As Single  (read/write)`  
  Returns or sets the number of points that are cropped off the right side of the specified picture or OLE object. Read/write Single.
- `CropTop As Single  (read/write)`  
  Returns or sets the number of points that are cropped off the top of the specified picture or OLE object. Read/write Single.
- `TransparencyColor As MsoRGBType  (read/write)`  
  Returns or sets the transparent color for the specified picture as a red-green-blue (RGB) value. For this property to take effect, the TransparentBackground property must be set to True. Applies to bitmaps only. Read/write Long.
- `TransparentBackground As MsoTriState  (read/write)`  
  Use the TransparencyColor property to set the transparent color. Applies to bitmaps only. Read/write MsoTriState.
- `Crop As Crop  (read-only)`  
  Returns a Crop object that represents the cropping settings for the specified PictureFormat object. Read-only.

## Methods (2)

- `IncrementBrightness(Increment As Single)`  
  Changes the brightness of the picture by the specified amount. Use the Brightness property to set the absolute brightness of the picture.
    - `Increment As Single` (required): Specifies how much to change the value of the Brightness property for the picture. A positive value makes the picture brighter; a negative value makes the picture darker.
- `IncrementContrast(Increment As Single)`  
  Changes the contrast of the picture by the specified amount. Use the Contrast property to set the absolute contrast for the picture.
    - `Increment As Single` (required): Specifies how much to change the value of the Contrast property for the picture. A positive value increases the contrast; a negative value decreases the contrast.
