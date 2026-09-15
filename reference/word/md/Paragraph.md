# Paragraph

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020957-0000-0000-C000-000000000046}  

Represents a single paragraph in a selection, range, or document. The Paragraph object is a member of the Paragraphs collection. The Paragraphs collection includes all the paragraphs in a selection, range, or document.

**Remarks:** Use Paragraphs (Index), where Index is the index number, to return a single Paragraph object. The following example right aligns the first paragraph in the active document. Use the Add, InsertParagraph, InsertParagraphAfter, or InsertParagraphBefore method to add a new, blank paragraph to a document. The following example adds a paragraph mark before the first paragraph in the selection. The following example also adds a paragraph mark before the first paragraph in the selection.

## Properties (49)

- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that is contained within the specified paragraph.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Paragraph object.
- `Format As ParagraphFormat  (read/write)`  
  Returns or sets a ParagraphFormat object that represents the formatting of the specified paragraph or paragraphs.
- `TabStops As TabStops  (read/write)`  
  Returns or sets a TabStops collection that represents all the custom tab stops for the specified paragraph. Read/write.
- `Borders As Borders  (read/write)`  
  Returns a Borders collection that represents all the borders for the specified paragraph.
- `DropCap As DropCap  (read-only)`  
  Returns a DropCap object that represents a dropped capital letter for the specified paragraph. Read-only.
- `Style As Variant  (read/write)`  
  Returns or sets the style for the specified object. Read/write Variant.
- `Alignment As WdParagraphAlignment  (read/write)`  
  Returns or sets a WdParagraphAlignment constant that represents the alignment for the specified paragraphs. Read/write.
- `KeepTogether As Long  (read/write)`  
  True if all lines in the specified paragraph remain on the same page when Microsoft Word repaginates the document. Read/write Long.
- `KeepWithNext As Long  (read/write)`  
  True if the specified paragraph remains on the same page as the paragraph that follows it when Microsoft Word repaginates the document. Read/write Long.
- `PageBreakBefore As Long  (read/write)`  
  True if a page break is forced before the specified paragraphs. Read/write Long.
- `NoLineNumber As Long  (read/write)`  
  True if line numbers are repressed for the specified paragraph. Read/write Long.
- `RightIndent As Single  (read/write)`  
  Returns or sets the right indent (in points) for the specified paragraph. Read/write Single.
- `LeftIndent As Single  (read/write)`  
  Returns or sets a Single that represents the left indent value (in points) for the specified paragraph. Read/write.
- `FirstLineIndent As Single  (read/write)`  
  Returns or sets the value (in points) for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent. Read/write Single.
- `LineSpacing As Single  (read/write)`  
  Returns or sets the line spacing (in points) for the specified paragraphs. Read/write Single.
- `LineSpacingRule As WdLineSpacing  (read/write)`  
  Returns or sets the line spacing for the specified paragraph. Read/write WdLineSpacing.
- `SpaceBefore As Single  (read/write)`  
  Returns or sets the spacing (in points) before the specified paragraphs. Read/write Single.
- `SpaceAfter As Single  (read/write)`  
  Returns or sets the amount of spacing (in points) after the specified paragraph or text column. Read/write Single.
- `Hyphenation As Long  (read/write)`  
  True if the specified paragraphs are included in automatic hyphenation. False if the specified paragraphs are to be excluded from automatic hyphenation. Read/write Long.
- `WidowControl As Long  (read/write)`  
  True if the first and last lines in the specified paragraph remain on the same page as the rest of the paragraph when Word repaginates the document. Read/write Long.
- `Shading As Shading  (read-only)`  
  Returns a Shading object that refers to the shading formatting for the specified paragraph.
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
- `OutlineLevel As WdOutlineLevel  (read/write)`  
  Returns or sets the outline level for the specified paragraph. Read/write WdOutlineLevel.
- `CharacterUnitRightIndent As Single  (read/write)`  
  Returns or sets the right indent value (in characters) for the specified paragraphs. Read/write Single.
- `CharacterUnitLeftIndent As Single  (read/write)`  
  Returns or sets the left indent value (in characters) for the specified paragraphs. Read/write Single.
- `CharacterUnitFirstLineIndent As Single  (read/write)`  
  Returns or sets the value (in characters) for a first-line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent. Read/write Single.
- `LineUnitBefore As Single  (read/write)`  
  Returns or sets the amount of spacing (in gridlines) before the specified paragraph. Read/write Single.
- `LineUnitAfter As Single  (read/write)`  
  Returns or sets the amount of spacing (in gridlines) after the specified paragraph. Read/write Single.
- `ReadingOrder As WdReadingOrder  (read/write)`  
  Returns or sets the reading order of the specified paragraph without changing the alignment. Read/write WdReadingOrder.
- `ID As String  (read/write)`  
  Returns or sets the identifying label for the specified object when the current document is saved as a webpage. Read/write String.
- `SpaceBeforeAuto As Long  (read/write)`  
  True if Microsoft Word automatically sets the amount of spacing before the specified paragraphs. Read/write Long.
- `SpaceAfterAuto As Long  (read/write)`  
  True if Microsoft Word automatically sets the amount of spacing after the specified paragraphs. Read/write Long.
- `IsStyleSeparator As Boolean  (read-only)`  
  True if a paragraph contains a special hidden paragraph mark that allows Microsoft Word to appear to join paragraphs of different paragraph styles. Read-only Boolean.
- `MirrorIndents As Long  (read/write)`  
  Returns or sets a Long that represents whether left and right indents are the same width. Can be True, False, or wdUndefined. Read/write.
- `TextboxTightWrap As WdTextboxTightWrap  (read/write)`  
  Returns or sets a WdTextboxTightWrap constant that represents how tightly text wraps around shapes or text boxes. Read/write.
- `ListNumberOriginal As Integer  (read-only)`  
  Returns an Integer that represents the original list level for a paragraph. Read-only.
- `CollapsedState As Boolean  (read/write)`  
  Returns or sets whether the specified paragraph is currently in a collapsed state. Read/write Boolean.
- `CollapseHeadingByDefault As Boolean  (read/write)`  
  Returns or sets whether the specified paragraph is collapsed by default when the document loads. Read/write Boolean.

## Methods (23)

- `CloseUp()`  
  Removes any spacing before the specified paragraph.
- `OpenUp()`  
  Sets spacing before the specified paragraphs to 12 points.
- `OpenOrCloseUp()`  
  Toggles the spacing before a paragraph.
- `TabHangingIndent(Count As Integer)`  
  Sets a hanging indent to a specified number of tab stops. .
    - `Count As Integer` (required): The number of tab stops to indent (if positive) or the number of tab stops to remove from the indent (if negative).
- `TabIndent(Count As Integer)`  
  Sets the left indent for the specified paragraphs to a specified number of tab stops. .
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
  Indents a paragraphs by a specified number of characters.
    - `Count As Integer` (required): The number of characters by which the specified paragraphs are to be indented.
- `IndentFirstLineCharWidth(Count As Integer)`  
  Indents the first line of one or more paragraphs by a specified number of characters.
    - `Count As Integer` (required): The number of characters by which the first line of each specified paragraph is to be indented.
- `Next([Count As Variant]) As Paragraph`  
  Returns a Paragraph object that represents the next paragraph.
    - `Count As Variant` (optional): The number of paragraphs by which you want to move ahead. The default value is one.
- `Previous([Count As Variant]) As Paragraph`  
  Returns the previous paragraph as a Paragraph object.
    - `Count As Variant` (optional): The number of paragraphs by which you want to move back. The default value is 1.
- `OutlinePromote()`  
  Applies the previous heading level style (Heading 1 through Heading 8) to the specified paragraph or paragraphs.
- `OutlineDemote()`  
  Applies the next heading level style (Heading 1 through Heading 8) to the specified paragraph or paragraphs.
- `OutlineDemoteToBody()`  
  Demotes the specified paragraph to body text by applying the Normal style.
- `Indent()`  
  Indents one or more paragraphs by one level.
- `Outdent()`  
  Removes one level of indent for one or more paragraphs.
- `SelectNumber()`  
  Selects the number or bullet in a list.
- `ListAdvanceTo([Level1 As Integer], [Level2 As Integer], [Level3 As Integer], [Level4 As Integer], [Level5 As Integer], [Level6 As Integer], [Level7 As Integer], [Level8 As Integer], [Level9 As Integer])`  
  Sets the list levels for a paragraph in a list.
- `ResetAdvanceTo()`  
  Resets a paragraph that uses custom list levels to the original level settings.
- `SeparateList()`  
  Separates a list into two separate lists. For numbered lists, the new list restarts numbering at the starting number, usually 1.
- `JoinList()`  
  Joins a list paragraph with the closest list above or below the specified paragraph.
