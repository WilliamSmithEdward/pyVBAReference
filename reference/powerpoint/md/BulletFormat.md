# BulletFormat

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493497-5A91-11CF-8700-00AA0060263B}  

Represents bullet formatting.

**Example:**

```vba
With ActivePresentation.Slides(1).Shapes(2)

    With .TextFrame.TextRange.ParagraphFormat.Bullet

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

## Properties (11)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Character As Long  (read/write)`  
  Returns or sets the Unicode character value that is used for bullets in the specified text. Read/write.
- `RelativeSize As Single  (read/write)`  
  Returns or sets the bullet size relative to the size of the first text character in the paragraph. Read/write.
- `UseTextColor As MsoTriState  (read/write)`  
  Determines whether the specified bullets are set to the color of the first text character in the paragraph. Read/write.
- `UseTextFont As MsoTriState  (read/write)`  
  Determines whether the specified bullets are set to the font of the first text character in the paragraph. Read/write.
- `Font As Font  (read-only)`  
  Returns a Font object that represents character formatting. Read-only.
- `Type As PpBulletType  (read/write)`  
  Represents the type of bullet. Read/write.
- `Style As PpNumberedBulletStyle  (read/write)`  
  Returns or sets the bullet style. Read/write.
- `StartValue As Long  (read/write)`  
  Returns or sets the beginning value of a bulleted list when the Type property of the BulletFormat object is set to ppBulletNumbered. Read/write.
- `Number As Long  (read-only)`  
  Returns the bullet number of a paragraph when the Type property of the BulletFormat object is set to ppBulletNumbered. Read-only.

## Methods (1)

- `Picture(Picture As String)`  
  Sets the graphics file to be used for bullets in a bulleted list when the Type property of the BulletFormat object is set to ppBulletPicture.
