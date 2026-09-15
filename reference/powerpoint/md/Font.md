# Font

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493495-5A91-11CF-8700-00AA0060263B}  

Represents character formatting for text or a bullet. The Font object is a member of the Fonts collection. The Fonts collection contains all the fonts used in a presentation.

**Example:**

```vba
With ActivePresentation.Slides(1).Shapes.Title _

        .TextFrame.TextRange

    .Text = "Volcano Coffee"

    With .Font

        .Italic = True

        .Name = "Palatino"

        .Color.RGB = RGB(0, 0, 255)

    End With

End With
```

## Properties (20)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Color As ColorFormat  (read-only)`  
  Returns or sets the color of the current Font object. Read/write.
- `Bold As MsoTriState  (read/write)`  
  Determines whether the character format is bold. Read/write.
- `Italic As MsoTriState  (read/write)`  
  True if the font or range is formatted as italic. Read/write Long.
- `Shadow As MsoTriState  (read/write)`  
  Determines whether the specified text has a shadow. Read/write.
- `Emboss As MsoTriState  (read/write)`  
  Determines whether the character format is embossed. Read/write.
- `Underline As MsoTriState  (read/write)`  
  Determines whether the specified text (for the Font object) or the font style (for the FontInfo object) is underlined. Read/write.
- `Subscript As MsoTriState  (read/write)`  
  Determines whether the specified text is subscript. Read/write.
- `Superscript As MsoTriState  (read/write)`  
  Determines whether the specified text is superscript. Read/write.
- `BaselineOffset As Single  (read/write)`  
  Returns or sets the baseline offset for the specified superscript or subscript characters. Read/write.
- `Embedded As MsoTriState  (read-only)`  
  Determines whether the specified font is embedded in the presentation. Read-only.
- `Embeddable As MsoTriState  (read-only)`  
  Determines whether the specified font can be embedded in the presentation. Read-only.
- `Size As Single  (read/write)`  
  Returns or sets the character size, in points. Read/write.
- `Name As String  (read/write)`  
  Returns or sets the name of the specified object. Read/write.
- `NameFarEast As String  (read/write)`  
  Returns or sets the Asian font name. Read/write.
- `NameAscii As String  (read/write)`  
  Returns or sets the font used for ASCII characters (characters with character set numbers within the range of 0 to 127). Read/write.
- `AutoRotateNumbers As MsoTriState  (read/write)`  
  Returns or sets lateral compression. Read/write.
- `NameOther As String  (read/write)`  
  Returns or sets the font used for characters whose character set numbers are greater than 127. Read/write.
- `NameComplexScript As String  (read/write)`  
  Returns or sets the complex script font name. Used for mixed language text. Read/write.
