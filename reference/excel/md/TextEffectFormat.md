# TextEffectFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000C031F-0000-0000-C000-000000000046}  

Contains properties and methods that apply to WordArt objects.

**Remarks:** Use the TextEffect property of the Shape object to return a TextEffectFormat object.

**Example:**

```vba
Set myDocument = Worksheets(1)
With myDocument.Shapes(1).TextEffect
 .FontName = "Courier New"
 .FontBold = True
 .FontItalic = True
End With
```

## Properties (15)

- `Application As Object  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Alignment As MsoTextEffectAlignment  (read/write)`  
  Returns or sets an MsoTextEffectAlignment value that represents the alignment for WordArt.
- `FontBold As MsoTriState  (read/write)`  
  Returns msoTrue if the font in the specified WordArt is bold. Read/write MsoTriState.
- `FontItalic As MsoTriState  (read/write)`  
  Returns msoTrue if the font in the specified WordArt is italic. Read/write MsoTriState.
- `FontName As String  (read/write)`  
  Returns or sets the name of the font in the specified WordArt. Read/write String.
- `FontSize As Single  (read/write)`  
  Returns or sets the font size for the specified WordArt, in points. Read/write Single.
- `KernedPairs As MsoTriState  (read/write)`  
  Returns msoTrue if character pairs in the specified WordArt are kerned. Read/write MsoTriState.
- `NormalizedHeight As MsoTriState  (read/write)`  
  Returns msoTrue if all characters (both uppercase and lowercase) in the specified WordArt are the same height. Read/write MsoTriState.
- `PresetShape As MsoPresetTextEffectShape  (read/write)`  
  Returns or sets the shape of the specified WordArt. Read/write MsoPresetTextEffectShape.
- `PresetTextEffect As MsoPresetTextEffect  (read/write)`  
  Returns or sets the style of the specified WordArt. Read/write MsoPresetTextEffect.
- `RotatedChars As MsoTriState  (read/write)`  
  Returns msoTrue if characters in the specified WordArt are rotated 90 degrees relative to the WordArt's bounding shape. Returns msoFalse if characters in the specified WordArt retain their original orientation relative to the bounding shape. Read/write MsoTriState.
- `Text As String  (read/write)`  
  Returns or sets the text for the specified object. Read/write String.
- `Tracking As Single  (read/write)`  
  Returns or sets the ratio of the horizontal space allotted to each character in the specified WordArt to the width of the character. Can be a value from 0 (zero) through 5. Large values for this property specify ample space between characters; values less than 1 can produce character overlap. Read/write Single.

## Methods (1)

- `ToggleVerticalText()`  
  Switches the text flow in the specified WordArt from horizontal to vertical, or vice versa.
