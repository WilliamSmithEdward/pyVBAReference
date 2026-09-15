# Paragraphs

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020958-0000-0000-C000-000000000046}  

A collection of Paragraph objects in a selection, range, or document.

**Remarks:** Use the Paragraphs property to return the Paragraphs collection. The following example formats the selected paragraphs to be double-spaced and right-aligned. Use the Add, InsertParagraph, InsertParagraphAfter, or InsertParagraphBefore method to add a new paragraph to a document. The following example adds a new paragraph before the first paragraph in the selection. The following example also adds a paragraph before the first paragraph in the selection. Use Paragraphs (Index), where Index is the index number, to return a single Paragraph object. The following example right aligns the first paragraph in the active document. The Count property for this collection in a document returns the number of items in the main story only. To count items in other stories use the collection with the Range object.

## Properties (44)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of paragraphs in the collection. Read-only.
- `First As Paragraph  (read-only)`  
  Returns a Paragraph object that represents the first item in the Paragraphs collection.
- `Last As Paragraph  (read-only)`  
  Returns a Paragraph object that represents the last item in the collection of paragraphs.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Paragraphs object.
- `Format As ParagraphFormat  (read/write)`  
  Returns or sets a ParagraphFormat object that represents the formatting of the specified paragraph or paragraphs.
- `TabStops As TabStops  (read/write)`  
  Returns or sets a TabStops collection that represents all the custom tab stops for the specified paragraphs. Read/write.
- `Borders As Borders  (read/write)`  
  Returns a Borders collection that represents all the borders for the specified object.
- `Style As Variant  (read/write)`  
  Returns or sets the style for the specified paragraphs. Read/write Variant.
- `Alignment As WdParagraphAlignment  (read/write)`  
  Returns or sets a WdParagraphAlignment constant that represents the alignment for the specified paragraphs. Read/write.
- `KeepTogether As Long  (read/write)`  
  True if all lines in the specified paragraphs remain on the same page when Microsoft Word repaginates the document. Read/write Long.
- `KeepWithNext As Long  (read/write)`  
  True if the specified paragraphs remain on the same page as the paragraphs that follow it when Microsoft Word repaginates the document. Read/write Long.
- `PageBreakBefore As Long  (read/write)`  
  True if a page break is forced before the specified paragraphs. Can be True, False, or wdUndefined. Read/write Long.
- `NoLineNumber As Long  (read/write)`  
  True if line numbers are repressed for the specified paragraphs. Can be True, False, or wdUndefined. Read/write Long.
- `RightIndent As Single  (read/write)`  
  Returns or sets the right indent (in points) for the specified paragraphs. Read/write Single.
- `LeftIndent As Single  (read/write)`  
  Returns or sets a Single that represents the left indent value (in points) for the specified paragraphs. Read/write.
- `FirstLineIndent As Single  (read/write)`  
  Returns or sets the value (in points) for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent. Read/write Single.
- `LineSpacing As Single  (read/write)`  
  Returns or sets the line spacing (in points) for the specified paragraphs. Read/write Single.
- `LineSpacingRule As WdLineSpacing  (read/write)`  
  Returns or sets the line spacing for the specified paragraphs. Read/write WdLineSpacing.
- `SpaceBefore As Single  (read/write)`  
  Returns or sets the spacing (in points) before the specified paragraphs. Read/write Single.
- `SpaceAfter As Single  (read/write)`  
  Returns or sets the amount of spacing (in points) after the specified paragraph or text column. Read/write Single.
- `Hyphenation As Long  (read/write)`  
  True if the specified paragraphs are included in automatic hyphenation. False if the specified paragraphs are to be excluded from automatic hyphenation. Read/write Long.
- `WidowControl As Long  (read/write)`  
  True if the first and last lines in the specified paragraph remain on the same page as the rest of the paragraph when Word repaginates the document. Can be True, False or wdUndefined. Read/write Long.
- `Shading As Shading  (read-only)`  
  Returns a Shading object that refers to the shading formatting for the specified paragraphs.
- `FarEastLineBreakControl As Long  (read/write)`  
  True if Microsoft Word applies East Asian line-breaking rules to the specified paragraphs. Returns wdUndefined if the FarEastLineBreakControl property is set to True for only some of the specified paragraphs. Read/write Long.
- `WordWrap As Long  (read/write)`  
  True if Microsoft Word wraps Latin text in the middle of a word in the specified paragraphs. Read/write Long.
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

## Methods (20)

- `Item(Index As Long) As Paragraph`  
  Returns an individual Paragraph object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
- `Add([Range As Variant]) As Paragraph`  
  Returns a Paragraph object that represents a new, blank paragraph added to the document. When the caller is not the last paragraph in a document, Paragraph adds a new, blank paragraph at the insertion point. When adding a paragraph to the very end, the insertion point will be before the last paragraph, and thus the new paragraph will be inserted as the second-to-last.
    - `Range As Variant` (optional): The range before which you want the new paragraph to be added. The new paragraph doesn't replace the range.
- `CloseUp()`  
  Removes any spacing before the specified paragraphs.
- `OpenUp()`  
  Sets spacing before the specified paragraphs to 12 points.
- `OpenOrCloseUp()`  
  Toggles spacing before paragraphs.
- `TabHangingIndent(Count As Integer)`  
  Sets a hanging indent to a specified number of tab stops.
    - `Count As Integer` (required): The number of tab stops to indent (if positive) or the number of tab stops to remove from the indent (if negative).
- `TabIndent(Count As Integer)`  
  Sets the left indent for the specified paragraphs to a specified number of tab stops.
    - `Count As Integer` (required): The number of tab stops to indent (if positive) or the number of tab stops to remove from the indent (if negative).
- `Reset()`  
  Removes manual paragraph formatting (formatting not applied using a style). .
- `Space1()`  
  Single-spaces the specified paragraphs.
- `Space15()`  
  Formats the specified paragraphs with 1.5-line spacing.
- `Space2()`  
  Double-spaces the specified paragraphs. .
- `IndentCharWidth(Count As Integer)`  
  Indents one or more paragraphs by a specified number of characters.
    - `Count As Integer` (required): The number of characters by which the specified paragraphs are to be indented.
- `IndentFirstLineCharWidth(Count As Integer)`  
  Indents the first line of one or more paragraphs by a specified number of characters.
    - `Count As Integer` (required): The number of characters by which the first line of each specified paragraph is to be indented.
- `OutlinePromote()`  
  Applies the previous heading level style (Heading 1 through Heading 8) to the specified paragraph or paragraphs.
- `OutlineDemote()`  
  Applies the next heading level style (Heading 1 through Heading 8) to the specified paragraphs.
- `OutlineDemoteToBody()`  
  Demotes the specified paragraph or paragraphs to body text by applying the Normal style.
- `Indent()`  
  Indents one or more paragraphs by one level.
- `Outdent()`  
  Removes one level of indent for one or more paragraphs.
- `IncreaseSpacing()`  
  Increases the spacing before and after paragraphs in six-point increments.
- `DecreaseSpacing()`  
  Decreases the spacing before and after paragraphs in six-point increments.
