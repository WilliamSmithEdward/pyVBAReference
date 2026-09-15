# ParagraphFormat

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493496-5A91-11CF-8700-00AA0060263B}  

Represents the paragraph formatting of a text range.

**Example:**

```vba
ActivePresentation.Slides(1).Shapes(2).TextFrame.TextRange _

    .ParagraphFormat.Alignment = ppAlignLeft
```

## Properties (15)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Alignment As PpParagraphAlignment  (read/write)`  
  Returns or sets the alignment for each paragraph in the specified paragraph format. Read/write.
- `Bullet As BulletFormat  (read-only)`  
  Returns a BulletFormat object that represents bullet formatting for the specified paragraph format. Read-only.
- `LineRuleBefore As MsoTriState  (read/write)`  
  Determines whether line spacing before the first line in each paragraph is set to a specific number of points or lines. Read/write.
- `LineRuleAfter As MsoTriState  (read/write)`  
  Determines whether line spacing after the last line in each paragraph is set to a specific number of points or lines. Read/write.
- `LineRuleWithin As MsoTriState  (read/write)`  
  Determines whether line spacing between base lines is set to a specific number of points or lines. Read/write.
- `SpaceBefore As Single  (read/write)`  
  Returns or sets the amount of space before the first line in each paragraph of the specified text, in points or lines. Read/write.
- `SpaceAfter As Single  (read/write)`  
  Returns or sets the amount of space after the last line in each paragraph of the specified text, in points or lines. Read/write.
- `SpaceWithin As Single  (read/write)`  
  Returns or sets the amount of space between base lines in the specified text, in points or lines. Read/write.
- `BaseLineAlignment As PpBaselineAlignment  (read/write)`  
  Returns or sets the base line alignment for the specified paragraph. Read/write.
- `FarEastLineBreakControl As MsoTriState  (read/write)`  
  Returns or sets the line break control option if you have an Asian language setting specified. Read/write.
- `WordWrap As MsoTriState  (read/write)`  
  Used only with Kanji characters. Read/write.
- `HangingPunctuation As MsoTriState  (read/write)`  
  Returns or sets the hanging punctuation option if you have an Asian language setting specified. Read/write.
- `TextDirection As PpDirection  (read/write)`  
  Returns or sets the text direction for the specified paragraph. Read/write.
