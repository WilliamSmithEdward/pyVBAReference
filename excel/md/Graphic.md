# Graphic

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024459-0000-0000-C000-000000000046}  

Contains properties that apply to header and footer picture objects.

**Remarks:** Use the following properties of the PageSetup object to return the Graphic object: - CenterFooterPicture - CenterHeaderPicture - LeftFooterPicture - LeftHeaderPicture - RightFooterPicture - RightHeaderPicture

**Example:**

```vba
Sub InsertPicture()

 With ActiveSheet.PageSetup.LeftFooterPicture
 .FileName = "C:\Sample.jpg"
 .Height = 275.25
 .Width = 463.5
 .Brightness = 0.36
 .ColorType = msoPictureGrayscale
 .Contrast = 0.39
 .CropBottom = -14.4
 .CropLeft = -28.8
 .CropRight = -14.4
 .CropTop = 21.6
 End With

 ' Enable the image to show up in the left footer.
 ActiveSheet.PageSetup.LeftFooter = "&G"

End Sub
```

## Properties (14)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
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
- `Filename As String  (read/write)`  
  Returns or sets the URL (on the intranet or the web) or path (local or network) to the location where the specified source object was saved. Read/write String.
- `Height As Single  (read/write)`  
  Returns or sets a Single value that represents the height, in points, of the object.
- `LockAspectRatio As MsoTriState  (read/write)`  
  True if the specified shape retains its original proportions when you resize it. False if you can change the height and width of the shape independently of one another when you resize it. Read/write MsoTriState.
- `Width As Single  (read/write)`  
  Returns or sets a Single value that represents the width, in points, of the object.
