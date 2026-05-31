# Crop

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03D3-0000-0000-C000-000000000046}  

An object used to remove a portion of an image.

**Example:**

```vba
Sub CropImage()
 ActivePresentation.Slides(1).Shapes.AddPicture "c:\myImage.png", msoFalse, msoTrue, 250,150, 200, 200
 ActivePresentation.Slides(1).Shapes(1).PictureFormat.Crop.PictureHeight = 100
 ActivePresentation.Slides(1).Shapes(1).PictureFormat.Crop.PictureWidth = 100
 ActivePresentation.Slides(1).Shapes(1).PictureFormat.Crop.PictureOffsetX = 0
 ActivePresentation.Slides(1).Shapes(1).PictureFormat.Crop.PictureOffsetY = 0
 ActivePresentation.Slides(1).Shapes(1).PictureFormat.Crop.ShapeHeight = 100
 ActivePresentation.Slides(1).Shapes(1).PictureFormat.Crop.ShapeWidth = 100
 ActivePresentation.Slides(1).Shapes(1).PictureFormat.Crop.ShapeLeft = 330
 ActivePresentation.Slides(1).Shapes(1).PictureFormat.Crop.ShapeTop = 170
End Sub
```

## Properties (10)

- `Application As Object  (read-only)`  
  Gets the Application object of the host application. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the Crop object was created. Read-only.
- `PictureOffsetX As Single  (read/write)`  
  Gets or sets the _x_-axis offset of the image that is to be cropped. Read/write.
- `PictureOffsetY As Single  (read/write)`  
  Gets or sets the _y_-axis offset of the image that is to be cropped. Read/write.
- `PictureWidth As Single  (read/write)`  
  Gets or sets the width of the image that is to be cropped. Read/write.
- `PictureHeight As Single  (read/write)`  
  Gets or sets the height of the image that is to be cropped. Read/write.
- `ShapeLeft As Single  (read/write)`  
  Gets or sets the location of the left-side of a shape that is used to crop an image. Read/write.
- `ShapeTop As Single  (read/write)`  
  Gets or sets the location of the top of a shape that is used to crop an image. Read/write.
- `ShapeWidth As Single  (read/write)`  
  Gets or sets the width of a shape that is used to crop an image. Read/write.
- `ShapeHeight As Single  (read/write)`  
  Gets or sets the height of a shape that is used to crop an image. Read/write.
