# TextEffectFormat

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493482-5A91-11CF-8700-00AA0060263B}  

Contains properties and methods that apply to WordArt objects.

**Example:**

```vba
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(1).TextEffect

    .FontName = "Courier New"

    .FontBold = True

    .FontItalic = True

End With
```

## Properties (15)

- `Application As Object  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Creator As Long  (read-only)`  
  Returns a Long that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in Microsoft PowerPoint, this property returns the hexadecimal number 50575054. Read-only.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Alignment As MsoTextEffectAlignment  (read/write)`  
  Returns or sets the alignment for the specified WordArt. Read/write.
- `FontBold As MsoTriState  (read/write)`  
  Determines whether the font in the specified WordArt is bold. Read/write.
- `FontItalic As MsoTriState  (read/write)`  
  Determines whether the font in the specified WordArt is italic. Read/write.
- `FontName As String  (read/write)`  
  Returns or sets the name of the font in the specified WordArt. Read/write.
- `FontSize As Single  (read/write)`  
  Returns or sets the font size for the specified WordArt in points. Read/write.
- `KernedPairs As MsoTriState  (read/write)`  
  Determines whether the character pairs in the specified WordArt are kerned. Read/write.
- `NormalizedHeight As MsoTriState  (read/write)`  
  Determines whether the characters (both uppercase and lowercase) in the specified WordArt are the same height. Read/write.
- `PresetShape As MsoPresetTextEffectShape  (read/write)`  
  Returns or sets the shape of the specified WordArt. Read/write.
- `PresetTextEffect As MsoPresetTextEffect  (read/write)`  
  Returns or sets the style of the specified WordArt. Read/write.
- `RotatedChars As MsoTriState  (read/write)`  
  Determines whether characters in the specified WordArt are rotated 90 degrees relative to the WordArt's bounding shape. Read/write.
- `Text As String  (read/write)`  
  Returns or sets a String that represents the text contained in the specified object. Read/write.
- `Tracking As Single  (read/write)`  
  Returns or sets the ratio of the horizontal space allotted to each character in the specified text to the width of the character. Read/write.

## Methods (1)

- `ToggleVerticalText()`  
  Switches the text flow in the specified WordArt from horizontal to vertical, or vice versa.
