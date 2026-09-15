# Font

**Type:** Class  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209F5-0000-0000-C000-000000000046}  

Contains font attributes (such as font name, font size and color) for an object.

**Remarks:** Use the Font property to return the Font object. The following instruction applies bold formatting to the selection. The following example formats the first paragraph in the active document as 24point Arial and italic. The following example changes the formatting of the Heading 2 style in the active document to Arial and bold. Use the New keyword to create a new, stand-alone Font object. The following example creates a Font object, sets some formatting properties, and then applies the Font object to the first paragraph in the active document. You can also duplicate a Font object by using the Duplicate property. The following example creates a new character style with the character formatting from the selection and italic formatting. The formatting of the selection is not changed.

## Properties (51)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Font object.
- `Duplicate As Font  (read-only)`  
  Returns a copy of a Font object that represents the character formatting of the specified font.
- `Bold As Long  (read/write)`  
  True if the font is formatted as bold. Read/write Long.
- `Italic As Long  (read/write)`  
  True if the font or range is formatted as italic. Read/write Long.
- `Hidden As Long  (read/write)`  
  True if the font is formatted as hidden text. Read/write Long.
- `SmallCaps As Long  (read/write)`  
  True if the font is formatted as small capital letters. Read/write Long.
- `AllCaps As Long  (read/write)`  
  True if the font is formatted as all capital letters. Read/write Long.
- `StrikeThrough As Long  (read/write)`  
  True if the font is formatted as strikethrough text. Read/write Long.
- `DoubleStrikeThrough As Long  (read/write)`  
  True if the specified font is formatted as double strikethrough text. .
- `ColorIndex As WdColorIndex  (read/write)`  
  Returns or sets a WdColorIndex constant that represents the color for the specified font. Read/write .
- `Subscript As Long  (read/write)`  
  True if the font is formatted as subscript. Read/write Long.
- `Superscript As Long  (read/write)`  
  True if the font is formatted as superscript. Read/write Long.
- `Underline As WdUnderline  (read/write)`  
  Returns or sets the type of underline applied to the font. Read/write WdUnderline.
- `Size As Single  (read/write)`  
  Returns or sets the font size, in points. Read/write Single.
- `Name As String  (read/write)`  
  Returns or sets the name of the specified object. Read/write String.
- `Position As Long  (read/write)`  
  Returns or sets the position of text (in points) relative to the base line. Read/write Long.
- `Spacing As Single  (read/write)`  
  Returns or sets the spacing (in points) between characters. Read/write Single.
- `Scaling As Long  (read/write)`  
  Returns or sets the scaling percentage applied to the font. Read/write Long.
- `Shadow As Long  (read/write)`  
  True if the specified font is formatted as shadowed. Read/write Long.
- `Outline As Long  (read/write)`  
  True if the font is formatted as outline. Read/write Long.
- `Emboss As Long  (read/write)`  
  True if the specified font is formatted as embossed. Read/write Long.
- `Kerning As Single  (read/write)`  
  Returns or sets the minimum font size for which Microsoft Word will adjust kerning automatically. Read/write Single.
- `Engrave As Long  (read/write)`  
  True if the font is formatted as engraved. Read/write Long.
- `Borders As Borders  (read/write)`  
  Returns a Borders collection that represents all the borders for the specified font.
- `Shading As Shading  (read-only)`  
  Returns a Shading object that refers to the shading formatting for the specified font.
- `EmphasisMark As WdEmphasisMark  (read/write)`  
  Returns or sets a WdEmphasisMark constant that represents the emphasis mark for a character or designated character string. Read/write.
- `DisableCharacterSpaceGrid As Boolean  (read/write)`  
  True if Microsoft Word ignores the number of characters per line for the corresponding Font object. Read/write Boolean.
- `NameFarEast As String  (read/write)`  
  Returns or sets an East Asian font name. Read/write String.
- `NameAscii As String  (read/write)`  
  Returns or sets the font used for Latin text (characters with character codes from 0 (zero) through 127). Read/write String.
- `NameOther As String  (read/write)`  
  Returns or sets the font used for characters with character codes from 128 through 255. Read/write String.
- `BoldBi As Long  (read/write)`  
  True if the font is formatted as bold. Read/write Long.
- `ItalicBi As Long  (read/write)`  
  True if the font or range is formatted as italic. Read/write Long.
- `SizeBi As Single  (read/write)`  
  Returns or sets the font size in points. Read/write Single.
- `NameBi As String  (read/write)`  
  Returns or sets the name of the font in a right-to-left language document. Read/write String.
- `ColorIndexBi As WdColorIndex  (read/write)`  
  Returns or sets the color for the specified Font object in a right-to-left language document. Read/write WdColorIndex.
- `DiacriticColor As WdColor  (read/write)`  
  Returns or sets the 24-bit color to be used for diacritics for the specified Font object. Read/write.
- `UnderlineColor As WdColor  (read/write)`  
  Returns or sets the 24-bit color of the underline for the specified Font object. .
- `Glow As GlowFormat  (read/write)`  
  Returns a GlowFormat object that represents the glow formatting for the font used by the specified range of text. Read-only.
- `Reflection As ReflectionFormat  (read/write)`  
  Returns a ReflectionFormat object that represents the reflection formatting for a shape. Read-only.
- `TextShadow As ShadowFormat  (read/write)`  
  Returns a ShadowFormat object that specifies the shadow formatting for the specified font.
- `Fill As FillFormat  (read/write)`  
  Returns a FillFormat object that contains fill formatting properties for the font used by the specified range of text. Read-only.
- `Line As LineFormat  (read/write)`  
  Returns a LineFormat object that specifies the formatting for a line. Read/write.
- `ThreeD As ThreeDFormat  (read/write)`  
  Returns a ThreeDFormat object that contains 3D effect formatting properties for the specified font. Read-only.
- `TextColor As ColorFormat  (read-only)`  
  Returns a ColorFormat object that represents the color for the specified font. Read-only.
- `Ligatures As WdLigatures  (read/write)`  
  Returns or sets the ligatures setting for the specified Font object. Read/writeWdLigatures.
- `NumberForm As WdNumberForm  (read/write)`  
  Returns or sets the number form setting for an OpenType font. Read/write WdNumberForm.
- `NumberSpacing As WdNumberSpacing  (read/write)`  
  Returns or sets the number spacing setting for a font. Read/write WdNumberSpacing.
- `ContextualAlternates As Long  (read/write)`  
  Specifies whether or not contextual alternates are enabled for the specified font. Read/write Long.
- `StylisticSet As WdStylisticSet  (read/write)`  
  Specifies the stylistic set for the specified font. Read/write WdStylisticSet.

## Methods (4)

- `Grow()`  
  Increases the font size to the next available size.
- `Shrink()`  
  Decreases the font size to the next available size.
- `Reset()`  
  Removes manual character formatting (formatting not applied using a style). For example, if you manually format a word as bold and the underlying style is plain text (not bold), the Reset method removes the bold format.
- `SetAsTemplateDefault()`  
  Sets the specified font formatting as the default for the active document and all new documents based on the active template.
