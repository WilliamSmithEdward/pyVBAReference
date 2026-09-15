# PictureFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209CB-0000-0000-C000-000000000046}  

Contains properties and methods that apply to pictures and OLE objects. The LinkFormat object contains properties and methods that apply to linked OLE objects only. The OLEFormat object contains properties and methods that apply to OLE objects whether or not they're linked.

**Remarks:** Use the PictureFormat property to return a PictureFormat object. The following example sets the brightness, contrast, and color transformation for shape one on the active document and crops 18 points off the bottom of the shape. For this example to work, shape one must be either a picture or an OLE object.

## Properties (13)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified PictureFormat object.
- `Brightness As Single  (read/write)`  
  Returns or sets the brightness of the specified picture or OLE object. The value for this property must be a number from 0.0 (dimmest) to 1.0 (brightest). Read/write Single.
- `ColorType As MsoPictureColorType  (read/write)`  
  Returns or sets the type of color transformation applied to the specified picture or OLE object. Read/write MsoPictureColorType.
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
- `TransparencyColor As Long  (read/write)`  
  Returns or sets the transparent color for the specified picture as a red-green-blue (RGB) value. Read/write Long.
- `TransparentBackground As MsoTriState  (read/write)`  
  MsoTrue if the parts of the picture that are defined with a transparent color actually appear transparent. Use the TransparencyColor property to set the transparent color. Applies to bitmaps only. Read/write MsoTriState.
- `Crop As Crop  (read/write)`  
  Returns or sets a Crop object that represents an image cropping. Read/write.

## Methods (2)

- `IncrementBrightness(Increment As Single)`  
  Changes the brightness of the picture by the specified amount.
    - `Increment As Single` (required): Specifies how much to change the value of the Brightness property for the picture. A positive value makes the picture brighter; a negative value makes the picture darker.
- `IncrementContrast(Increment As Single)`  
  Changes the contrast of the picture by the specified amount.
    - `Increment As Single` (required): Specifies how much to change the value of the Contrast property for the picture. A positive value increases the contrast; a negative value decreases the contrast.
