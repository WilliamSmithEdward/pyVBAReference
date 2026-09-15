# TextEffectFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209CF-0000-0000-C000-000000000046}  

Contains properties and methods that apply to WordArt objects.

**Remarks:** Use the TextEffect property to return a TextEffectFormat object. The following example sets the font name and formatting for shape one on the active document. For this example to work, shape one must be a WordArt object.

## Properties (15)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified TextEffectFormat object.
- `Alignment As MsoTextEffectAlignment  (read/write)`  
  Returns or sets an MsoTextEffectAlignment constant that represents the alignment for the specified text effect. Read/write.
- `FontBold As MsoTriState  (read/write)`  
  Sets the font to bold for the specified Word Art shape. Read/write MsoTriState.
- `FontItalic As MsoTriState  (read/write)`  
  Italicizes WordArt text. Read/write MsoTriState.
- `FontName As String  (read/write)`  
  Returns or sets the name of the font for the dropped capital letter. Read/write String.
- `FontSize As Single  (read/write)`  
  Returns or sets the font size for the specified WordArt, in points. Read/write Single.
- `KernedPairs As MsoTriState  (read/write)`  
  Indicates that character pairs in a WordArt object have been kerned. Read/write MsoTriState.
- `NormalizedHeight As MsoTriState  (read/write)`  
  MsoTrue if all characters (both uppercase and lowercase) in the specified WordArt are the same height. Read/write MsoTriState.
- `PresetShape As MsoPresetTextEffectShape  (read/write)`  
  Returns or sets the shape of the specified WordArt. Read/write MsoPresetTextEffectShape.
- `PresetTextEffect As MsoPresetTextEffect  (read/write)`  
  Returns or sets the style of the specified WordArt. The values for this property correspond to the formats in the WordArt Gallery dialog box (Insert menu), numbered from left to right, top to bottom. Read/write MsoPresetTextEffect.
- `RotatedChars As MsoTriState  (read/write)`  
  MsoTrue if characters in the specified WordArt are rotated 90 degrees relative to the WordArt's bounding shape. MsoFalse if characters in the specified WordArt retain their original orientation relative to the bounding shape. Read/write MsoTriState.
- `Text As String  (read/write)`  
  Returns or sets the text in the specified object. Read/write String.
- `Tracking As Single  (read/write)`  
  Returns or sets the ratio of the horizontal space allotted to each character in the specified WordArt in relation to the width of the character. Read/write Single.

## Methods (1)

- `ToggleVerticalText()`  
  Switches the text flow in the specified WordArt from horizontal to vertical, or vice versa.
