# ParagraphFormat2

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0399-0000-0000-C000-000000000046}  

Represents the paragraph formatting of a text range.

**Example:**

```vba
ActivePresentation.Slides(1).Shapes(2).TextFrame2.TextRange2 _
 .ParagraphFormat2.Alignment = ppAlignLeft
```

## Properties (21)

- `Application As Object  (read-only)`  
  Gets an object that represents the application that contains the object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a value representing the application that created the ParagraphFormat2 object. Read-only.
- `Parent As Object  (read-only)`  
  Gets the parent object for the ParagraphFormat2 object. Read-only.
- `Alignment As MsoParagraphAlignment  (read/write)`  
  Gets or sets a value specifying the alignment of the paragraph. Read/write.
- `BaselineAlignment As MsoBaselineAlignment  (read/write)`  
  Gets or sets a constant that represents the vertical position of fonts in a paragraph. Read/write.
- `Bullet As BulletFormat2  (read-only)`  
  Gets a BulletFormat2 object for the paragraph. Read-only.
- `FarEastLineBreakLevel As MsoTriState  (read/write)`  
  Gets or sets the East Asian line break control level for the specified paragraph. Read/write.
- `FirstLineIndent As Single  (read/write)`  
  Gets or sets the value (in points) for a first line or hanging indent. Read/write.
- `HangingPunctuation As MsoTriState  (read/write)`  
  Determines whether hanging punctuation is enabled for the specified paragraphs. Read/write.
- `IndentLevel As Long  (read/write)`  
  Gets or sets a value representing the indent level assigned to text in the selected paragraph. Read/write.
- `LeftIndent As Single  (read/write)`  
  Gets or sets a value that represents the left indent value (in points) for the specified paragraphs. Read/write.
- `LineRuleAfter As MsoTriState  (read/write)`  
  Determines whether line spacing after the last line in each paragraph is set to a specific number of points or lines. Read/write.
- `LineRuleBefore As MsoTriState  (read/write)`  
  Determines whether line spacing before the first line in each paragraph is set to a specific number of points or lines. Read/write.
- `LineRuleWithin As MsoTriState  (read/write)`  
  Determines whether line spacing between base lines is set to a specific number of points or lines. Read/write.
- `RightIndent As Single  (read/write)`  
  Gets or sets the right indent (in points) for the specified paragraphs. Read/write.
- `SpaceAfter As Single  (read/write)`  
  Gets or sets the amount of spacing (in points) after the specified paragraph. Read/write.
- `SpaceBefore As Single  (read/write)`  
  Gets or sets the spacing (in points) before the specified paragraphs. Read/write.
- `SpaceWithin As Single  (read/write)`  
  Gets or sets the amount of space between base lines in the specified paragraph, in points or lines. Read/write.
- `TabStops As TabStops2  (read-only)`  
  Gets a TabStops2 collection that represents all the custom tab stops for the specified paragraphs. Read-only.
- `TextDirection As MsoTextDirection  (read/write)`  
  Gets or sets the text direction for the specified paragraph. Read/write.
- `WordWrap As MsoTriState  (read/write)`  
  Determines whether the application wraps the Latin text in the middle of a word in the specified paragraphs. Read/write.
