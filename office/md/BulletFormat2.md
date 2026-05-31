# BulletFormat2

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03B9-0000-0000-C000-000000000046}  

Represents bullet formatting.

**Example:**

```vba
With ActivePresentation.Slides(1).Shapes(2)
 With .TextFrame.TextRange.ParagraphFormat.BulletFormat2
 .Visible = True
 .RelativeSize = 1.25
 .Character = 169
 With .Font
 .Color.RGB = RGB(255, 255, 0)
 .Name = "Symbol"
 End With
 End With
End With
```

## Properties (13)

- `Application As Object  (read-only)`  
  Gets an object that represents the BulletFormat2 object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the BulletFormat2 object was created. Read-only.
- `Parent As Object  (read-only)`  
  Gets the parent of the BulletFormat2 object. Read-only.
- `Character As Long  (read/write)`  
  Gets or sets the Unicode character value that is used for bullets in the specified text. Read/write.
- `Font As Font2  (read-only)`  
  Gets a Font2 object that represents character formatting for a BulletFormat2 object. Read-only.
- `Number As Long  (read-only)`  
  Gets the bullet number of a paragraph. Read-only.
- `RelativeSize As Single  (read/write)`  
  Returns or sets the bullet size relative to the size of the first text character in the paragraph. Read/write.
- `StartValue As Long  (read/write)`  
  Gets or sets the beginning value of a bulleted list. Read/write.
- `Style As MsoNumberedBulletStyle  (read/write)`  
  Returns or sets a constant that represents the style of a bullet. Read/write.
- `Type As MsoBulletType  (read/write)`  
  Gets or sets a constant that represents the type of bullet. Read/write.
- `UseTextColor As MsoTriState  (read/write)`  
  Determines whether the specified bullets are set to the color of the first text character in the paragraph. Read/write.
- `UseTextFont As MsoTriState  (read/write)`  
  Determines whether the specified bullets are set to the font of the first text character in the paragraph. Read/write.
- `Visible As MsoTriState  (read/write)`  
  Gets or sets a value that specifies whether the bullet is visible. Read/write.

## Methods (1)

- `Picture(FileName As String)`  
  Sets the graphics file to be used for bullets in a bulleted list.
    - `FileName As String` (required): The file name of a valid graphics file.
