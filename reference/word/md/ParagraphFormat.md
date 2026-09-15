# ParagraphFormat

**Type:** Class  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209F4-0000-0000-C000-000000000046}  

Represents all the formatting for a paragraph.

**Remarks:** Use the Format property to return the ParagraphFormat object for a paragraph or paragraphs. The ParagraphFormat property returns the ParagraphFormat object for a selection, range, style, Find object, or Replacement object. The following example centers the third paragraph in the active document. The following example finds the next double-spaced paragraph after the selection. Use Visual Basic's New keyword to create a new, standalone ParagraphFormat object. The following example creates a ParagraphFormat object, sets some formatting properties for it, and then applies all of its properties to the first paragraph in the active document. You can also make a standalone copy of an existing ParagraphFormat object by using the Duplicate property. The following example duplicates the paragraph formatting of the first paragraph in the active document and stores the formatting in _myDup_. The example changes the left indent of _myDup_ to 1 inch, creates a new document, inserts text into the document, and applies the paragraph formatting of _myDup_ to the text.

## Properties (43)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ParagraphFormat object.
- `Duplicate As ParagraphFormat  (read-only)`  
  Returns a read-only ParagraphFormat object that represents the paragraph formatting of the specified paragraph.
- `Style As Variant  (read/write)`  
  Returns or sets the style for the specified object. Read/write Variant.
- `Alignment As WdParagraphAlignment  (read/write)`  
  Returns or sets a WdParagraphAlignment constant that represents the alignment for the specified paragraphs. Read/write.
- `KeepTogether As Long  (read/write)`  
  True if all lines in the specified paragraphs remain on the same page when Microsoft Word repaginates the document. Read/write Long.
- `KeepWithNext As Long  (read/write)`  
  True if the specified paragraph remains on the same page as the paragraph that follows it when Microsoft Word repaginates the document. Read/write Long.
- `PageBreakBefore As Long  (read/write)`  
  True if a page break is forced before the specified paragraphs. Can be True, False, or wdUndefined. Read/write Long.
- `NoLineNumber As Long  (read/write)`  
  True if line numbers are repressed for the specified paragraphs. Can be True, False, or wdUndefined. Read/write Long.
- `RightIndent As Single  (read/write)`  
  Returns or sets the right indent (in points) for the specified paragraphs. Read/write Single.
- `LeftIndent As Single  (read/write)`  
  Returns or sets a Single that represents the left indent value (in points) for the specified paragraph formatting. Read/write.
- `FirstLineIndent As Single  (read/write)`  
  Returns or sets the value (in points) for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent. Read/write Single.
- `LineSpacing As Single  (read/write)`  
  Returns or sets the line spacing (in points) for the specified paragraphs. Read/write Single.
- `LineSpacingRule As WdLineSpacing  (read/write)`  
  Returns or sets the line spacing for the specified paragraph formatting. Read/write WdLineSpacing.
- `SpaceBefore As Single  (read/write)`  
  Returns or sets the spacing (in points) before the specified paragraphs. Read/write Single.
- `SpaceAfter As Single  (read/write)`  
  Returns or sets the amount of spacing (in points) after the specified paragraph or text column. Read/write Single.
- `Hyphenation As Long  (read/write)`  
  True if the specified paragraphs are included in automatic hyphenation. False if the specified paragraphs are to be excluded from automatic hyphenation. Read/write Long.
- `WidowControl As Long  (read/write)`  
  True if the first and last lines in the specified paragraph remain on the same page as the rest of the paragraph when Word repaginates the document. Can be True, False or wdUndefined. Read/write Long.
- `FarEastLineBreakControl As Long  (read/write)`  
  True if Microsoft Word applies East Asian line-breaking rules to the specified paragraphs. Returns wdUndefined if the FarEastLineBreakControl property is set to True for only some of the specified paragraphs. Read/write Long.
- `WordWrap As Long  (read/write)`  
  True if Microsoft Word wraps Latin text in the middle of a word in the specified paragraphs or text frames. Read/write Long.
- `HangingPunctuation As Long  (read/write)`  
  True if hanging punctuation is enabled for the specified paragraphs. This property returns wdUndefined if it's set to True for only some of the specified paragraphs. Read/write Long.
- `HalfWidthPunctuationOnTopOfLine As Long  (read/write)`  
  True if Microsoft Word changes punctuation symbols at the beginning of a line to half-width characters for the specified paragraphs. This property returns wdUndefined if it's set to True for only some of the specified paragraphs. Read/write Long.
- `AddSpaceBetweenFarEastAndAlpha As Long  (read/write)`  
  True if Microsoft Word is set to automatically add spaces between Japanese and Latin text for the specified paragraphs. This property returns wdUndefined if it's set to True for only some of the specified paragraphs. Read/write Long.
- `AddSpaceBetweenFarEastAndDigit As Long  (read/write)`  
  True if Microsoft Word is set to automatically add spaces between Japanese text and numbers for the specified paragraphs. This property returns wdUndefined if it's set to True for only some of the specified paragraphs. Read/write Long.
- `BaseLineAlignment As WdBaselineAlignment  (read/write)`  
  Returns or sets a WdBaselineAlignment constant that represents the vertical position of fonts on a line. Read/write.
- `AutoAdjustRightIndent As Long  (read/write)`  
  True if Microsoft Word is set to automatically adjust the right indent for the specified paragraphs if you've specified a set number of characters per line. Returns wdUndefined if the AutoAdjustRightIndent property is set to True for only some of the specified paragraphs. Read/write Long.
- `DisableLineHeightGrid As Long  (read/write)`  
  True if Microsoft Word aligns characters in the specified paragraphs to the line grid when a set number of lines per page is specified. Returns wdUndefined if the DisableLineHeightGrid property is set to True for only some of the specified paragraphs. Read/write Long.
- `TabStops As TabStops  (read/write)`  
  Returns or sets a TabStops collection that represents all the custom tab stops for the specified paragraphs. Read/write.
- `Borders As Borders  (read/write)`  
  Returns a Borders collection that represents all the borders for the specified object.
- `Shading As Shading  (read-only)`  
  Returns a Shading object that refers to the shading formatting for the specified object.
- `OutlineLevel As WdOutlineLevel  (read/write)`  
  Returns or sets the outline level for the specified paragraphs. Read/write WdOutlineLevel.
- `CharacterUnitRightIndent As Single  (read/write)`  
  Returns or sets the right indent value (in characters) for the specified paragraphs. Read/write Single.
- `CharacterUnitLeftIndent As Single  (read/write)`  
  Returns or sets the left indent value (in characters) for the specified paragraphs. Read/write Single.
- `CharacterUnitFirstLineIndent As Single  (read/write)`  
  Returns or sets the value (in characters) for a first-line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent. Read/write Single.
- `LineUnitBefore As Single  (read/write)`  
  Returns or sets the amount of spacing (in gridlines) before the specified paragraphs. Read/write Single.
- `LineUnitAfter As Single  (read/write)`  
  Returns or sets the amount of spacing (in gridlines) after the specified paragraphs. Read/write Single.
- `ReadingOrder As WdReadingOrder  (read/write)`  
  Returns or sets the reading order of the specified paragraphs without changing their alignment. Read/write WdReadingOrder.
- `SpaceBeforeAuto As Long  (read/write)`  
  True if Microsoft Word automatically sets the amount of spacing before the specified paragraphs. Read/write Long.
- `SpaceAfterAuto As Long  (read/write)`  
  True if Microsoft Word automatically sets the amount of spacing after the specified paragraphs. Read/write Long.
- `MirrorIndents As Long  (read/write)`  
  Returns or sets a Long that represents whether left and right indents are the same width. Can be True, False, or wdUndefined. Read/write.
- `TextboxTightWrap As WdTextboxTightWrap  (read/write)`  
  Returns or sets a WdTextboxTightWrap constant that represents how tightly text wraps around shapes or text boxes. Read/write.
- `CollapsedByDefault As Long  (read/write)`  
  Returns or sets whether the specified paragraph format is collapsed by default. Read/write Long.

## Methods (11)

- `CloseUp()`  
  Removes any spacing before paragraphs in the specified paragraph format.
- `OpenUp()`  
  Sets spacing before the specified paragraphs to 12 points.
- `OpenOrCloseUp()`  
  Toggles the spacing before the specified paragraphs.
- `TabHangingIndent(Count As Integer)`  
  Sets a hanging indent to a specified number of tab stops. .
    - `Count As Integer` (required): The number of tab stops to indent (if positive) or the number of tab stops to remove from the indent (if negative).
- `TabIndent(Count As Integer)`  
  Sets the left indent for the specified paragraphs to a specified number of tab stops.
    - `Count As Integer` (required): The number of tab stops to indent (if positive) or the number of tab stops to remove from the indent (if negative).
- `Reset()`  
  Removes manual paragraph formatting (formatting not applied using a style).
- `Space1()`  
  Single-spaces the specified paragraphs.
- `Space15()`  
  Formats the specified paragraphs with 1.5-line spacing.
- `Space2()`  
  Double-spaces the specified paragraphs.
- `IndentCharWidth(Count As Integer)`  
  Indents one or more paragraphs by a specified number of characters.
    - `Count As Integer` (required): The number of characters by which the specified paragraphs are to be indented.
- `IndentFirstLineCharWidth(Count As Integer)`  
  Indents the first line of one or more paragraphs by a specified number of characters.
    - `Count As Integer` (required): The number of characters by which the first line of each specified paragraph is to be indented.
