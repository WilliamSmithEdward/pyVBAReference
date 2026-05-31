# Font2

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C039A-0000-0000-C000-000000000046}  

Contains font attributes (for example, font name, font size, and color) for an object.

**Example:**

```vba
With ActiveDocument.Styles(wdStyleHeading2).Font2
 .Name = "Arial"
 .Italic = True
End With
```

## Properties (36)

- `Application As Object  (read-only)`  
  Gets an object that represents the application that the Font2 object is used in. Read-only.
- `Creator As Long  (read-only)`  
  Gets a value indicating the application that the object was created in. Read-only.
- `Parent As Object  (read-only)`  
  Gets the parent of the Font2 object. Read-only.
- `Bold As MsoTriState  (read/write)`  
  Gets or sets a value specifying whether the font should be bold. Read/write.
- `Italic As MsoTriState  (read/write)`  
  Gets or sets a value specifying whether the text for a selection is italic. Read/write.
- `Strike As MsoTextStrike  (read/write)`  
  Gets or sets a value specifying the strike format used for a selection of text. Read/write.
- `Caps As MsoTextCaps  (read/write)`  
  Gets or sets a value specifying that the text should be capitalized. Read/write.
- `AutorotateNumbers As MsoTriState  (read/write)`  
  Gets or sets a value that specifies whether the numbers in a numbered list should be rotated when the text is rotated. Read/write.
- `BaselineOffset As Single  (read/write)`  
  Gets or sets a value specifying the horizontal offset of the selected font. Read/write.
- `Kerning As Single  (read/write)`  
  Gets or sets a value specifying the amount of spacing between text characters. Read/write.
- `Size As Single  (read/write)`  
  Gets or sets a value specifying the size of the font. Read/write.
- `Spacing As Single  (read/write)`  
  Gets or sets a value specifying the spacing between characters in a selection of text. Read/write.
- `UnderlineStyle As MsoTextUnderlineType  (read/write)`  
  Gets or sets a value specifying the underline style for the selected text. Read/write.
- `Allcaps As MsoTriState  (read/write)`  
  True if the font is formatted as all capital letters. Read/write.
- `DoubleStrikeThrough As MsoTriState  (read/write)`  
  True if the specified font is formatted as double strikethrough text. Read/write.
- `Equalize As MsoTriState  (read/write)`  
  Gets or sets a value specifying whether the text for a selection should be spaced equal distances apart. Read/write.
- `Fill As FillFormat  (read-only)`  
  Gets the formatting properties for the font of the specified text. Read-only.
- `Glow As GlowFormat  (read-only)`  
  Gets a value indicating whether the font is displayed as a glow effect. Read-only.
- `Reflection As ReflectionFormat  (read-only)`  
  Gets a value specifying the type of reflection format for the selection of text. Read-only.
- `Line As LineFormat  (read-only)`  
  Gets a value specifying the format of a line. Read-only.
- `Shadow As ShadowFormat  (read-only)`  
  Gets the value specifying the type of shadow effect for the selection of text. Read-only.
- `Highlight As ColorFormat  (read-only)`  
  Gets a value indicating whether the font is displayed as highlighted. Read-only.
- `UnderlineColor As ColorFormat  (read-only)`  
  Gets a value specifying the color of the underline for the selected text. Read-only.
- `Smallcaps As MsoTriState  (read/write)`  
  Gets or sets a value specifying whether small caps should be used with the selection of text. Small caps are the same height as the lowercase letters in a selection of text. Read/write.
- `SoftEdgeFormat As MsoSoftEdgeType  (read/write)`  
  Gets or sets a value specifying the type of soft edge effect used in a selection of text. Read/write.
- `StrikeThrough As MsoTriState  (read/write)`  
  Gets or sets a value specifying that the text should be rendered in a strikethrough appearance. Read/write.
- `Subscript As MsoTriState  (read/write)`  
  Gets or sets a value specifying that the selected text should be displayed as subscript. Read/write.
- `Superscript As MsoTriState  (read/write)`  
  Gets or sets a value specifying that the selected text should be displayed as superscript. Read/write.
- `WordArtformat As MsoPresetTextEffect  (read/write)`  
  Gets or sets a value specifying the text effect for the selected text. Read/write.
- `Embeddable As MsoTriState  (read-only)`  
  Gets a value indicating whether the font can be embedded on a page. Read-only.
- `Embedded As MsoTriState  (read-only)`  
  Gets a value specifying whether the font is embedded in a page. Read-only.
- `Name As String  (read/write)`  
  Gets or sets a value specifying the font to use for a selection. Read/write.
- `NameAscii As String  (read/write)`  
  Gets or sets the font used for Latin text&mdash;characters with character codes from 0 (zero) through 127. Read/write.
- `NameComplexScript As String  (read/write)`  
  Gets or sets the complex script font name. Used for mixed language text. Read/write.
- `NameFarEast As String  (read/write)`  
  Gets or sets an East Asian font name. Read/write.
- `NameOther As String  (read/write)`  
  Gets or sets the font used for characters whose character set numbers are greater than 127. Read/write.
