# Range

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002095E-0000-0000-C000-000000000046}  

Represents a contiguous area in a document. Each Range object is defined by a starting and ending character position.

**Remarks:** Similar to the way bookmarks are used in a document, Range objects are used in Visual Basic procedures to identify specific portions of a document. However, unlike a bookmark, a Range object only exists while the procedure that defined it is running. Range objects are independent of the selection. That is, you can define and manipulate a range without changing the selection. You can also define multiple ranges in a document, while there can be only one selection per pane. Use the Range method to return a Range object defined by the given starting and ending character positions. The following example returns a Range object that refers to the first 10 characters in the active document. Use the Range property to return a Range object defined by the beginning and end of another object. The Range property applies to many objects (for example, Paragraph, Bookmark, and Cell). The following example returns a Range object that refers to the first paragraph in the active document. The following example returns a Range object that refers to the second through fourth paragraphs in the active document. For more information about working with Range objects, see Working with Range objects.

## Properties (96)

- `Text As String  (read/write)`  
  Returns or sets the text in the specified range or selection. Read/write String. Read/write String.
- `FormattedText As Range  (read/write)`  
  Returns or sets a Range object that includes the formatted text in the specified range or selection. Read/write.
- `Start As Long  (read/write)`  
  Returns or sets the starting character position of a range. Read/write Long.
- `End As Long  (read/write)`  
  Returns or sets the ending character position of a range. Read/write Long.
- `Font As Font  (read/write)`  
  Returns or sets a Font object that represents the character formatting of the specified object. Read/write Font.
- `Duplicate As Range  (read-only)`  
  Returns a read-only Range object that represents all the properties of the specified range.
- `StoryType As WdStoryType  (read-only)`  
  Returns the story type for the specified range, selection, or bookmark. Read-only WdStoryType.
- `Tables As Tables  (read-only)`  
  Returns a Tables collection that represents all the tables in the specified range. Read-only.
- `Words As Words  (read-only)`  
  Returns a Words collection that represents all the words in a range. Read-only.
- `Sentences As Sentences  (read-only)`  
  Returns a Sentences collection that represents all the sentences in the range. Read-only.
- `Characters As Characters  (read-only)`  
  Returns a Characters collection that represents the characters in a range. Read-only.
- `Footnotes As Footnotes  (read-only)`  
  Returns a Footnotes collection that represents all the footnotes in a range. Read-only.
- `Endnotes As Endnotes  (read-only)`  
  Returns an Endnotes collection that represents all the endnotes in a range. Read-only.
- `Comments As Comments  (read-only)`  
  Returns a Comments collection that represents all the comments in the specified document, selection, or range. Read-only.
- `Cells As Cells  (read-only)`  
  Returns a Cells collection that represents the table cells in a range. Read-only.
- `Sections As Sections  (read-only)`  
  Returns a Sections collection that represents the sections in the specified range. Read-only.
- `Paragraphs As Paragraphs  (read-only)`  
  Returns a Paragraphs collection that represents all the paragraphs in the specified range. Read-only.
- `Borders As Borders  (read/write)`  
  Returns a Borders collection that represents all the borders for the specified object.
- `Shading As Shading  (read-only)`  
  Returns a Shading object that refers to the shading formatting for the specified object.
- `TextRetrievalMode As TextRetrievalMode  (read/write)`  
  Returns a TextRetrievalMode object that controls how text is retrieved from the specified Range. Read/write.
- `Fields As Fields  (read-only)`  
  Returns a Fields collection that represents all the fields in the range. Read-only.
- `FormFields As FormFields  (read-only)`  
  Returns a FormFields collection that represents all the form fields in the range. Read-only.
- `Frames As Frames  (read-only)`  
  Returns a Frames collection that represents all the frames in a range. Read-only.
- `ParagraphFormat As ParagraphFormat  (read/write)`  
  Returns or sets a ParagraphFormat object that represents the paragraph settings for the specified range. Read/write.
- `ListFormat As ListFormat  (read-only)`  
  Returns a ListFormat object that represents all the list formatting characteristics of a range. Read-only.
- `Bookmarks As Bookmarks  (read-only)`  
  Returns a Bookmarks collection that represents all the bookmarks in a document, range, or selection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Range object.
- `Bold As Long  (read/write)`  
  True if the range is formatted as bold. Read/write Long.
- `Italic As Long  (read/write)`  
  True if the font or range is formatted as italic. Read/write Long.
- `Underline As WdUnderline  (read/write)`  
  Returns or sets the type of underline applied to a range. Read/write WdUnderline.
- `EmphasisMark As WdEmphasisMark  (read/write)`  
  Returns or sets the emphasis mark for a character or designated character string. Read/write WdEmphasisMark.
- `DisableCharacterSpaceGrid As Boolean  (read/write)`  
  True if Microsoft Word ignores the number of characters per line for the corresponding Range object. Read/write Boolean.
- `Revisions As Revisions  (read-only)`  
  Returns a Revisions collection that represents the tracked changes in the range. Read-only.
- `Style As Variant  (read/write)`  
  Returns or sets the style for the specified object. Read/write Variant.
- `StoryLength As Long  (read-only)`  
  Returns the number of characters in the story that contains the specified range. Read-only Long.
- `LanguageID As WdLanguageID  (read/write)`  
  Returns or sets a WdLanguageID constant that represents the language for the specified range. Read/write.
- `SynonymInfo As SynonymInfo  (read-only)`  
  Returns a SynonymInfo object that contains information from the thesaurus on synonyms, antonyms, or related words and expressions for the contents of a range.
- `Hyperlinks As Hyperlinks  (read-only)`  
  Returns a Hyperlinks collection that represents all the hyperlinks in the specified range. Read-only.
- `ListParagraphs As ListParagraphs  (read-only)`  
  Returns a ListParagraphs collection that represents all the numbered paragraphs in the range. Read-only.
- `Subdocuments As Subdocuments  (read-only)`  
  Returns a Subdocuments collection that represents all the subdocuments in the specified range or document. Read-only.
- `GrammarChecked As Boolean  (read/write)`  
  True if a grammar check has been run on the specified range or document. Read/write Boolean.
- `SpellingChecked As Boolean  (read/write)`  
  True if spelling has been checked throughout the specified range or document. False if all or some of the range or document has not been checked for spelling. Read/write Boolean.
- `HighlightColorIndex As WdColorIndex  (read/write)`  
  Returns or sets the highlight color for the specified range. Read/write WdColorIndex.
- `Columns As Columns  (read-only)`  
  Returns a Columns collection that represents all the table columns in the range. Read-only.
- `Rows As Rows  (read-only)`  
  Returns a Rows collection that represents all the table rows in a range. Read-only.
- `IsEndOfRowMark As Boolean  (read-only)`  
  True if the specified range is collapsed and is located at the end-of-row mark in a table. Read-only Boolean.
- `BookmarkID As Long  (read-only)`  
  Returns the number of the bookmark that encloses the beginning of the specified range; returns 0 (zero) if there is no corresponding bookmark. Read-only Long.
- `PreviousBookmarkID As Long  (read-only)`  
  Returns the number of the last bookmark that starts before or at the same place as the specified range. Read-only Long.
- `Find As Find  (read-only)`  
  Returns a Find object that contains the criteria for a find operation. Read-only.
- `PageSetup As PageSetup  (read/write)`  
  Returns a PageSetup object that's associated with the specified range.
- `ShapeRange As ShapeRange  (read-only)`  
  Returns a ShapeRange collection that represents all the Shape objects in the specified range. Read-only.
- `Case As WdCharacterCase  (read/write)`  
  Returns or sets a WdCharacterCase constant that represents the case of the text in the specified range. Read/write.
- `Information As Variant  (read-only)`  
  Returns information about the specified range. Read-only Variant.
- `ReadabilityStatistics As ReadabilityStatistics  (read-only)`  
  Returns a ReadabilityStatistics collection that represents the readability statistics for the specified document or range. Read-only.
- `GrammaticalErrors As ProofreadingErrors  (read-only)`  
  Returns a ProofreadingErrors collection that represents the sentences that failed the grammar check on the specified document or range. Read-only.
- `SpellingErrors As ProofreadingErrors  (read-only)`  
  Returns a ProofreadingErrors collection that represents the words identified as spelling errors in the specified range. Read-only.
- `Orientation As WdTextOrientation  (read/write)`  
  Returns or sets the orientation of text in a range when the Text Direction feature is enabled. Read/write WdTextOrientation.
- `InlineShapes As InlineShapes  (read-only)`  
  Returns an InlineShapes collection that represents all the InlineShape objects in a range. Read-only.
- `NextStoryRange As Range  (read-only)`  
  Returns a Range object that refers to the next story. Read-only Range.
- `LanguageIDFarEast As WdLanguageID  (read/write)`  
  Returns or sets an East Asian language for the specified object. Read/write WdLanguageID.
- `LanguageIDOther As WdLanguageID  (read/write)`  
  Returns or sets the language for the specified range. Read/write WdLanguageID.
- `LanguageDetected As Boolean  (read/write)`  
  Returns or sets a value that specifies whether Microsoft Word has detected the language of the specified text. Read/write Boolean.
- `FitTextWidth As Single  (read/write)`  
  Returns or sets the width (in the current measurement units) in which Microsoft Word fits the text in the current selection or range. Read/write Single.
- `HorizontalInVertical As WdHorizontalInVerticalType  (read/write)`  
  Returns or sets the formatting for horizontal text set within vertical text. Read/write WdHorizontalInVerticalType.
- `TwoLinesInOne As WdTwoLinesInOneType  (read/write)`  
  Returns or sets whether Microsoft Word sets two lines of text in one and specifies the characters that enclose the text, if any. Read/write WdTwoLinesInOneType.
- `CombineCharacters As Boolean  (read/write)`  
  True if the specified range contains combined characters. Read/write Boolean.
- `NoProofing As Long  (read/write)`  
  True if the spelling and grammar checker ignores the specified text. Read/write Long.
- `TopLevelTables As Tables  (read-only)`  
  Returns a Tables collection that represents the tables at the outermost nesting level in the current range. Read-only.
- `Scripts As Scripts  (read-only)`  
  Returns a Scripts collection that represents the collection of HTML scripts in the specified object.
- `CharacterWidth As WdCharacterWidth  (read/write)`  
  Returns or sets the character width of the specified range. Read/write WdCharacterWidth.
- `Kana As WdKana  (read/write)`  
  Returns or sets whether the specified range of Japanese language text is hiragana or katakana. Read/write WdKana.
- `BoldBi As Long  (read/write)`  
  True if the font or range is formatted as bold. Returns True, False, or wdUndefined (for a mixture of bold and non-bold text). Can be set to True, False, or wdToggle. Read/write Long.
- `ItalicBi As Long  (read/write)`  
  True if the font or range is formatted as italic. Read/write Long.
- `ID As String  (read/write)`  
  Returns or sets the identification name for the specified range. Read/write String.
- `HTMLDivisions As HTMLDivisions  (read-only)`  
  Returns an HTMLDivisions object that represents an HTML division in a web document.
- `ShowAll As Boolean  (read/write)`  
  True if all nonprinting characters (such as hidden text, tab marks, space marks, and paragraph marks) are displayed. Read/write Boolean.
- `Document As Document  (read-only)`  
  Returns a Document object associated with the specified range. Read-only.
- `FootnoteOptions As FootnoteOptions  (read-only)`  
  Returns FootnoteOptions object that represents the footnotes in a selection or range.
- `EndnoteOptions As EndnoteOptions  (read-only)`  
  Returns an EndnoteOptions object that represents the endnotes in a range.
- `Editors As Editors  (read-only)`  
  Returns an Editors object that represents all the users authorized to modify a selection or range within a document.
- `XML As String  (read-only)`  
  Returns a String that represents the XML text in the specified object. .
- `EnhMetaFileBits As Variant  (read-only)`  
  Returns a Variant that represents a picture representation of how a range of text appears.
- `OMaths As OMaths  (read-only)`  
  Returns an OMaths collection that represents the OMath objects within the specified range. Read-only.
- `CharacterStyle As Variant  (read-only)`  
  Returns a Variant that represents the style used to format one or more characters. Read-only.
- `ParagraphStyle As Variant  (read-only)`  
  Returns a Variant that represents the style used to format a paragraph. Read-only.
- `ListStyle As Variant  (read-only)`  
  Returns a Variant that represents the style used to format a bulleted list or numbered list. Read-only.
- `TableStyle As Variant  (read-only)`  
  Returns a Variant that represents the style used to format a table. Read-only.
- `ContentControls As ContentControls  (read-only)`  
  Returns a ContentControls collection that represents the content controls contained within a range. Read-only.
- `WordOpenXML As String  (read-only)`  
  Returns a String that represents the XML contained within the range in the Microsoft Word Open XML format. Read-only.
- `ParentContentControl As ContentControl  (read-only)`  
  Returns a ContentControl object that represents the parent content control for the specified range. Read-only.
- `Locks As CoAuthLocks  (read-only)`  
  Returns a CoAuthLocks collection object that represents all the locks in the range. Read-only.
- `Updates As CoAuthUpdates  (read-only)`  
  Returns a CoAuthUpdates collection object that represents all updates that were merged into the specified range at the last explicit save. Read-only.
- `Conflicts As Conflicts  (read-only)`  
  Returns a Conflicts collection object that contains all the conflict objects in the range. Read-only.
- `TextVisibleOnScreen As Long  (read-only)`  
  Returns a Long that indicates whether the text in the specified range is visible on the screen. Read-only.

## Methods (77)

- `Select()`  
  Selects the specified range.
- `SetRange(Start As Long, End As Long)`  
  Sets the starting and ending character positions for an existing range.
    - `Start As Long` (required): The starting character position of the range.
    - `End As Long` (required): The ending character position of the range.
- `Collapse([Direction As Variant])`  
  Collapses a range or selection to the starting or ending position. After a range or selection is collapsed, the starting and ending points are equal.
    - `Direction As Variant` (optional): The direction in which to collapse the range or selection. Can be either of the following WdCollapseDirection constants: wdCollapseEnd or wdCollapseStart. The default value is wdCollapseStart.
- `InsertBefore(Text As String)`  
  Inserts the specified text before the specified range.
    - `Text As String` (required): The text to be inserted.
- `InsertAfter(Text As String)`  
  Inserts the specified text at the end of a range.
    - `Text As String` (required): The text to be inserted.
- `Next([Unit As Variant], [Count As Variant]) As Range`  
  Returns a Range object that represents the specified unit relative to the specified range.
    - `Unit As Variant` (optional): The type of units by which to count. Can be any WdUnits constant.
    - `Count As Variant` (optional): The number of units by which you want to move ahead. The default value is one.
- `Previous([Unit As Variant], [Count As Variant]) As Range`  
  Returns the previous range a relative to the specified range.
    - `Unit As Variant` (optional): The type of units by which to count. Can be any WdUnits constant.
    - `Count As Variant` (optional): The number of units by which you want to move back. The default value is 1.
- `StartOf([Unit As Variant], [Extend As Variant]) As Long`  
  Moves or extends the start position of the specified range or selection to the beginning of the nearest specified text unit. This method returns a Long that indicates the number of characters by which the range or selection was moved or extended. The method returns a negative number if the movement is backward through the document.
    - `Unit As Variant` (optional): The unit by which the start position of the specified range or selection is to be moved. Can be any WdUnits constant except wdLine. The default value is wdWord.
    - `Extend As Variant` (optional): Specifies whether to move or extend the start of the range. If you use wdMove, both ends of the range or selection are moved to the beginning of the specified unit. If you use wdExtend, the beginning of the range or selection is extended to the beginning of the specified unit. The default value is wdMove.
- `EndOf([Unit As Variant], [Extend As Variant]) As Long`  
  Moves or extends the ending character position of a range to the end of the nearest specified text unit.
    - `Unit As Variant` (optional): The unit by which to move the ending character position. Can be any WdUnits, except wdLine. The default value is wdWord.
    - `Extend As Variant` (optional): Specifies whether to move or extend the end of the range. If the value is wdMove, both ends of the range or selection object are moved to the end of the specified unit. If wdExtend is used, the end of the range or selection is extended to the end of the specified unit. The default value is wdMove.
- `Move([Unit As Variant], [Count As Variant]) As Long`  
  Collapses the specified range to its start or end position and then moves the collapsed object by the specified number of units.
    - `Unit As Variant` (optional): The unit by which to move the range.
    - `Count As Variant` (optional): The number of units by which the specified range is to be moved. If Count is a positive number, the object is collapsed to its end position and moved backward in the document by the specified number of units. If Count is a negative number, the object is collapsed to its start position and moved forward by the specified number of units. The default value is 1. You can also control the collapse direction by using the Collapse method before using the Move method. If the range is in the middle of a unit or isn't collapsed, moving it to the beginning or end of the unit counts as moving it one full unit.
- `MoveStart([Unit As Variant], [Count As Variant]) As Long`  
  Moves the start position of the specified range.
    - `Unit As Variant` (optional): The unit by which start position of the specified range is to be moved.
    - `Count As Variant` (optional): The maximum number of units by which the specified range is to be moved. If Count is a positive number, the start position of the range is moved forward in the document. If it is a negative number, the start position is moved backward. If the start position is moved forward to a position beyond the end position, the range is collapsed and both the start and end positions are moved together. The default value is 1.
- `MoveEnd([Unit As Variant], [Count As Variant]) As Long`  
  Moves the ending character position of a range. .
    - `Unit As Variant` (optional): The unit by which to move the ending character position.
    - `Count As Variant` (optional): The number of units to move. If this number is positive, the ending character position is moved forward in the document. If this number is negative, the end is moved backward. If the ending position overtakes the starting position, the range collapses and both character positions move together. The default value is 1.
- `MoveWhile(Cset As Variant, [Count As Variant]) As Long`  
  Moves the specified range while any of the specified characters are found in the document.
    - `Cset As Variant` (required): One or more characters. This argument is case-sensitive.
    - `Count As Variant` (optional): The maximum number of characters by which the specified range is to be moved. Can be a number or either the wdForward or wdBackward constant. If Count is a positive number, the specified range is moved forward in the document, beginning at the end position. If it is a negative number, the range is moved backward, beginning at the start position. The default value is wdForward.
- `MoveStartWhile(Cset As Variant, [Count As Variant]) As Long`  
  Moves the start position of the specified range while any of the specified characters are found in the document.
    - `Cset As Variant` (required): One or more characters. This argument is case-sensitive.
    - `Count As Variant` (optional): The maximum number of characters by which the specified range is to be moved. Can be a number or either the wdForward or wdBackward constant. If Count is a positive number, the range is moved forward in the document. If it is a negative number, the range is moved backward. The default value is wdForward.
- `MoveEndWhile(Cset As Variant, [Count As Variant]) As Long`  
  Moves the ending character position of a range while any of the specified characters are found in the document.
    - `Cset As Variant` (required): One or more characters. This argument is case-sensitive.
    - `Count As Variant` (optional): The maximum number of characters by which the range is to be moved. Can be a number or either the wdForward or wdBackward constant. If Count is a positive number, the range is moved forward in the document. If it is a negative number, the range is moved backward. The default value is wdForward.
- `MoveUntil(Cset As Variant, [Count As Variant]) As Long`  
  Moves the specified range until one of the specified characters is found in the document.
    - `Cset As Variant` (required): One or more characters. If any character in Cset is found before the Count value expires, the specified range is positioned as an insertion point immediately before that character. This argument is case-sensitive.
    - `Count As Variant` (optional): The maximum number of characters by which the specified range is to be moved. Can be a number or either the wdForward or wdBackward constant. If Count is a positive number, the range is moved forward in the document, beginning at the end position. If it is a negative number, the range is moved backward, beginning at the start position. The default value is wdForward.
- `MoveStartUntil(Cset As Variant, [Count As Variant]) As Long`  
  Moves the start position of the specified range until one of the specified characters is found in the document.
    - `Cset As Variant` (required): One or more characters. This argument is case-sensitive.
    - `Count As Variant` (optional): The maximum number of characters by which the specified range is to be moved. Can be a number or either the wdForward or wdBackward constant. If Count is a positive number, the range is moved forward in the document. If it is a negative number, the range is moved backward. The default value is wdForward.
- `MoveEndUntil(Cset As Variant, [Count As Variant]) As Long`  
  Moves the end position of the specified range until any of the specified characters are found in the document. If the movement is forward in the document, the range is expanded.
    - `Cset As Variant` (required): One or more characters. This argument is case-sensitive.
    - `Count As Variant` (optional): The maximum number of characters by which the specified range is to be moved. Can be a number or either the wdForward or wdBackward constant. If Count is a positive number, the range is moved forward in the document. If it is a negative number, the range is moved backward. The default value is wdForward.
- `Cut()`  
  Removes the specified object from the document and places it on the Clipboard.
- `Copy()`  
  Copies the specified range to the Clipboard.
- `Paste()`  
  Inserts the contents of the Clipboard at the specified range.
- `InsertBreak([Type As Variant])`  
  Inserts a page, column, or section break.
    - `Type As Variant` (optional): The type of break to be inserted.Can be one of the WdBreakType constants. If omitted, the default value is wdPageBreak.
- `InsertFile(FileName As String, [Range As Variant], [ConfirmConversions As Variant], [Link As Variant], [Attachment As Variant])`  
  Inserts all or part of the specified file.
    - `FileName As String` (required): The path and file name of the file to be inserted. If you don't specify a path, Word assumes the file is in the current folder.
    - `Range As Variant` (optional): If the specified file is a Word document, this parameter refers to a bookmark. If the file is another type (for example, a Microsoft Excel worksheet), this parameter refers to a named range or a cell range (for example, R1C1:R3C4).
    - `ConfirmConversions As Variant` (optional): True to have Word prompt you to confirm conversion when inserting files in formats other than the Word Document format.
    - `Link As Variant` (optional): True to insert the file by using an INCLUDETEXT field.
    - `Attachment As Variant` (optional): True to insert the file as an attachment to an email message.
- `InStory(Range As Range) As Boolean`  
  True if the range to which this method is applied is in the same story as the range specified by the Range argument.
    - `Range As Range` (required): Specifies the range that this method uses to determine if it is contained within the specified Range object.
- `InRange(Range As Range) As Boolean`  
  Returns True if the range to which the method is applied is contained in the range specified by the Range argument.
    - `Range As Range` (required): Specifies the range that this method uses to determine if it is contained within the specified Range object.
- `Delete([Unit As Variant], [Count As Variant]) As Long`  
  Deletes the specified number of characters or words.
    - `Unit As Variant` (optional): The unit by which the collapsed range is to be deleted. Can be one of the WdUnits constants.
    - `Count As Variant` (optional): The number of units to be deleted. To delete units after the range, collapse the range and use a positive number. To delete units before the range, collapse the range and use a negative number.
- `WholeStory()`  
  Expands a range to include the entire story.
- `Expand([Unit As Variant]) As Long`  
  Expands the specified range or selection. Returns the number of characters added to the range or selection. Long.
    - `Unit As Variant` (optional): The unit by which to expand the range. Can be one of the following WdUnits constants: wdCharacter, wdWord, wdSentence, wdParagraph, wdSection, wdStory, wdCell wdColumn, wdRow, or wdTable.
- `InsertParagraph()`  
  Replaces the specified range with a new paragraph.
- `InsertParagraphAfter()`  
  Inserts a paragraph mark after a range.
- `InsertSymbol(CharacterNumber As Long, [Font As Variant], [Unicode As Variant], [Bias As Variant])`  
  Inserts a symbol in place of the specified range.
    - `CharacterNumber As Long` (required): The character number for the specified symbol. This value will always be the sum of 31 and the number that corresponds to the position of the symbol in the table of symbols (counting from left to right). For example, to specify a delta character at position 37 in the table of symbols in the Symbol font, set CharacterNumber to 68.
    - `Font As Variant` (optional): The name of the font that contains the symbol.
    - `Unicode As Variant` (optional): True to insert the unicode character specified by CharacterNumber; False to insert the ANSI character specified by CharacterNumber. The default value is False.
    - `Bias As Variant` (optional): Sets the font bias for symbols. This argument is useful for setting the correct font bias for East Asian characters. Can be one of the WdFontBias constants. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
- `CopyAsPicture()`  
  The CopyAsPicture method works the same way as the Copy method.
- `SortAscending()`  
  Sorts paragraphs or table rows in ascending alphanumeric order.
- `SortDescending()`  
  Sorts paragraphs in descending alphanumeric order.
- `IsEqual(Range As Range) As Boolean`  
  True if the range to which this method is applied is equal to the range specified by the Range argument.
    - `Range As Range` (required): The range to compare with the Range object defined by expression.
- `Calculate() As Single`  
  Calculates a mathematical expression within a range or selection. Returns the result as a Single.
- `GoTo([What As Variant], [Which As Variant], [Count As Variant], [Name As Variant]) As Range`  
  Returns a Range object that represents the start position of the specified item, such as a page, bookmark, or field.
    - `What As Variant` (optional): The kind of item to which the range is moved. Can be one of the wdGoToItem constants.
    - `Which As Variant` (optional): The item to which the range is moved. Can be one of the wdGoToDirection constants.
    - `Count As Variant` (optional): The number of the item in the document. The default value is 1. Only positive values are valid. To specify an item that precedes the range, use wdGoToPrevious as the Which argument and specify a Count value.
    - `Name As Variant` (optional): If the What argument is wdGoToBookmark, wdGoToComment, wdGoToField, or wdGoToObject, this argument specifies a name. Only positive values are valid. To specify an item that precedes the range, use wdGoToPrevious as the Which argument and specify a Count value.
- `GoToNext(What As WdGoToItem) As Range`  
  Returns a Range object that refers to the start position of the next item or location specified by the What argument. .
    - `What As WdGoToItem` (required): The item to where the specified range or selection is to be moved.
- `GoToPrevious(What As WdGoToItem) As Range`  
  Returns a Range object that refers to the start position of the previous item or location specified by the What argument.
    - `What As WdGoToItem` (required): The item to where the specified range or selection is to be moved.
- `PasteSpecial([IconIndex As Variant], [Link As Variant], [Placement As Variant], [DisplayAsIcon As Variant], [DataType As Variant], [IconFileName As Variant], [IconLabel As Variant])`  
  Inserts the contents of the Clipboard. .
    - `IconIndex As Variant` (optional): If DisplayAsIcon is True, this argument is a number that corresponds to the icon you want to use in the program file specified by IconFilename. Icons appear in the Change Icon dialog box: 0 (zero) corresponds to the first icon, 1 corresponds to the second icon, and so on. If this argument is omitted, the first (default) icon is used.
    - `Link As Variant` (optional): True to create a link to the source file of the Clipboard contents. The default value is False.
    - `Placement As Variant` (optional): Can be either of the following WdOLEPlacement constants: wdFloatOverText or wdInLine. The default value is wdInLine.
    - `DisplayAsIcon As Variant` (optional): True to display the link as an icon. The default value is False.
    - `DataType As Variant` (optional): A format for the Clipboard contents when they're inserted into the document. Can be any WdPasteDataType constant.
    - `IconFileName As Variant` (optional): If DisplayAsIcon is True, this argument is the path and file name for the file in which the icon to be displayed is stored.
    - `IconLabel As Variant` (optional): If DisplayAsIcon is True, this argument is the text that appears below the icon.
- `LookupNameProperties()`  
  Looks up a name in the global address book list and displays the Properties dialog box, which includes information about the specified name.
- `ComputeStatistics(Statistic As WdStatistic) As Long`  
  Returns a Long that represents a statistic based on the contents of the specified range.
    - `Statistic As WdStatistic` (required): The type of statistic to compute.
- `Relocate(Direction As Long)`  
  In outline view, moves the paragraphs within the specified range after the next visible paragraph or before the previous visible paragraph.
    - `Direction As Long` (required): The direction of the move.
- `CheckSynonyms()`  
  Displays the Thesaurus dialog box, which lists alternative word choices, or synonyms, for the text in the specified range.
- `InsertAutoText()`  
  Attempts to match the text in the specified range or the text surrounding the range with an existing AutoText entry name.
- `InsertDatabase([Format As Variant], [Style As Variant], [LinkToSource As Variant], [Connection As Variant], [SQLStatement As Variant], [SQLStatement1 As Variant], [PasswordDocument As Variant], [PasswordTemplate As Variant], [WritePasswordDocument As Variant], [WritePasswordTemplate As Variant], [DataSource As Variant], [From As Variant], [To As Variant], [IncludeFields As Variant])`  
  Retrieves data from a data source (for example, a separate Microsoft Word document, a Microsoft Excel worksheet, or a Microsoft Access database) and inserts the data as a table in place of the specified range.
    - `Format As Variant` (optional): A format listed in the Formats box in the Table AutoFormat dialog box (Table menu). Can be any of the WdTableFormat constants. A border is applied to the cells in the table by default.
    - `Style As Variant` (optional): The attributes of the AutoFormat specified by Format that are applied to the table.
    - `LinkToSource As Variant` (optional): True to establish a link between the new table and the data source.
    - `Connection As Variant` (optional): A range within which to perform the query specified by SQLStatement.
    - `SQLStatement As Variant` (optional): An optional query string that retrieves a subset of the data in a primary data source to be inserted into the document.
    - `SQLStatement1 As Variant` (optional): If the query string is longer than 255 characters, SQLStatement denotes the first portion of the string and SQLStatement1 denotes the second portion.
    - `PasswordDocument As Variant` (optional): The password (if any) required to open the data source. (See Remarks below.)
    - `PasswordTemplate As Variant` (optional): If the data source is a Word document, this argument is the password (if any) required to open the attached template. (See Remarks below.)
    - `WritePasswordDocument As Variant` (optional): The password required to save changes to the document. (See Remarks below.)
    - `WritePasswordTemplate As Variant` (optional): The password required to save changes to the template. (See Remarks below.)
    - `DataSource As Variant` (optional): The path and file name of the data source.
    - `From As Variant` (optional): The number of the first record in the range of records to be inserted.
    - `To As Variant` (optional): The number of the last record in the range of records to be inserted.
    - `IncludeFields As Variant` (optional): True to include field names from the data source in the first row of the new table.
- `AutoFormat()`  
  Automatically formats a document. Use the Kind property to specify a document type.
- `CheckGrammar()`  
  Begins a spelling and grammar check for the specified range.
- `CheckSpelling([CustomDictionary As Variant], [IgnoreUppercase As Variant], [AlwaysSuggest As Variant], [CustomDictionary2 As Variant], [CustomDictionary3 As Variant], [CustomDictionary4 As Variant], [CustomDictionary5 As Variant], [CustomDictionary6 As Variant], [CustomDictionary7 As Variant], [CustomDictionary8 As Variant], [CustomDictionary9 As Variant], [CustomDictionary10 As Variant])`  
  Begins a spelling check for the specified document or range.
    - `CustomDictionary As Variant` (optional): Either an expression that returns a Dictionary object or the file name of the custom dictionary.
    - `IgnoreUppercase As Variant` (optional): True if capitalization is ignored. If this argument is omitted, the current value of the IgnoreUppercase property is used.
    - `AlwaysSuggest As Variant` (optional): True for Microsoft Word to always suggest alternative spellings. If this argument is omitted, the current value of the SuggestSpellingCorrections property is used.
    - `CustomDictionary2 As Variant` (optional): Either an expression that returns a Dictionary object or the file name of the custom dictionary. You can specify as many as nine additional dictionaries.
    - `CustomDictionary3 As Variant` (optional): Either an expression that returns a Dictionary object or the file name of the custom dictionary. You can specify as many as nine additional dictionaries.
    - `CustomDictionary4 As Variant` (optional): Either an expression that returns a Dictionary object or the file name of the custom dictionary. You can specify as many as nine additional dictionaries.
    - `CustomDictionary5 As Variant` (optional): Either an expression that returns a Dictionary object or the file name of the custom dictionary. You can specify as many as nine additional dictionaries.
    - `CustomDictionary6 As Variant` (optional): Either an expression that returns a Dictionary object or the file name of the custom dictionary. You can specify as many as nine additional dictionaries.
    - `CustomDictionary7 As Variant` (optional): Either an expression that returns a Dictionary object or the file name of the custom dictionary. You can specify as many as nine additional dictionaries.
    - `CustomDictionary8 As Variant` (optional): Either an expression that returns a Dictionary object or the file name of the custom dictionary. You can specify as many as nine additional dictionaries.
    - `CustomDictionary9 As Variant` (optional): Either an expression that returns a Dictionary object or the file name of the custom dictionary. You can specify as many as nine additional dictionaries.
    - `CustomDictionary10 As Variant` (optional): Either an expression that returns a Dictionary object or the file name of the custom dictionary. You can specify as many as nine additional dictionaries.
- `GetSpellingSuggestions([CustomDictionary As Variant], [IgnoreUppercase As Variant], [MainDictionary As Variant], [SuggestionMode As Variant], [CustomDictionary2 As Variant], [CustomDictionary3 As Variant], [CustomDictionary4 As Variant], [CustomDictionary5 As Variant], [CustomDictionary6 As Variant], [CustomDictionary7 As Variant], [CustomDictionary8 As Variant], [CustomDictionary9 As Variant], [CustomDictionary10 As Variant]) As SpellingSuggestions`  
  Returns a SpellingSuggestions collection that represents the words suggested as spelling replacements for the first word in the specified range.
    - `CustomDictionary As Variant` (optional): Either an expression that returns a Dictionary object or the file name of the custom dictionary.
    - `IgnoreUppercase As Variant` (optional): True to ignore words in all uppercase letters. If this argument is omitted, the current value of the IgnoreUppercase property is used.
    - `MainDictionary As Variant` (optional): Either an expression that returns a Dictionary object or the file name of the main dictionary. If you don't specify a main dictionary, Microsoft Word uses the main dictionary that corresponds to the language formatting of the first word in the range.
    - `SuggestionMode As Variant` (optional): Specifies the way Word makes spelling suggestions. Can be one of the following WdSpellingWordType constants. The default value is wdSpellword.
- `InsertParagraphBefore()`  
  Inserts a new paragraph before the specified range.
- `NextSubdocument()`  
  Moves the range to the next subdocument.
- `PreviousSubdocument()`  
  Moves the range to the previous subdocument.
- `ConvertHangulAndHanja([ConversionsMode As Variant], [FastConversion As Variant], [CheckHangulEnding As Variant], [EnableRecentOrdering As Variant], [CustomDictionary As Variant])`  
  Converts the specified range from hangul to hanja or vice versa.
    - `ConversionsMode As Variant` (optional): Sets the direction for the conversion between hangul and hanja. Can be either of the following WdMultipleWordConversionsMode constants: wdHangulToHanja or wdHanjaToHangul. The default value is the current value of the MultipleWordConversionsMode property.
    - `FastConversion As Variant` (optional): True if Microsoft Word automatically converts a word with only one suggestion for conversion. The default value is the current value of the HangulHanjaFastConversion property.
    - `CheckHangulEnding As Variant` (optional): True if Word automatically detects hangul endings and ignores them. The default value is the current value of the CheckHangulEndings property. This argument is ignored if the ConversionsMode argument is set to wdHanjaToHangul.
    - `EnableRecentOrdering As Variant` (optional): True if Word displays the most recently used words at the top of the suggestions list. The default value is the current value of the EnableHangulHanjaRecentOrdering property.
    - `CustomDictionary As Variant` (optional): The name of a custom hangul-hanja conversion dictionary. Use this argument to use a custom dictionary with hangul-hanja conversions not contained in the main dictionary.
- `PasteAsNestedTable()`  
  Pastes a cell or group of cells as a nested table into the selected range.
- `ModifyEnclosure(Style As Variant, [Symbol As Variant], [EnclosedText As Variant])`  
  Adds, modifies, or removes an enclosure around the specified character or characters.
    - `Style As Variant` (required): The style of the enclosure. Can be any WdEncloseStyle constant.
    - `Symbol As Variant` (optional): The symbol in which to enclose the specified range. Can be any WdEnclosureType constant.
    - `EnclosedText As Variant` (optional): The characters that you want to enclose. If you include this argument, Microsoft Word replaces the specified range with the enclosed characters. If you don't specify text to enclose, Microsoft Word encloses all text in the specified range.
- `PhoneticGuide(Text As String, [Alignment As WdPhoneticGuideAlignmentType], [Raise As Long], [FontSize As Long], [FontName As String])`  
  Adds phonetic guides to the specified range.
    - `Text As String` (required): The phonetic text to add.
    - `Alignment As WdPhoneticGuideAlignmentType` (optional): The alignment of the added phonetic text.
    - `Raise As Long` (optional): The distance (in points) from the top of the text in the specified range to the top of the phonetic text. If no value is specified, Microsoft Word automatically sets the phonetic text at an optimum distance above the specified range.
    - `FontSize As Long` (optional): The font size to use for the phonetic text. If no value is specified, Word uses a font size 50 percent smaller than the text in the specified range.
    - `FontName As String` (optional): The name of the font to use for the phonetic text. If no value is specified, Word uses the same font as the text in the specified range.
- `InsertDateTime([DateTimeFormat As Variant], [InsertAsField As Variant], [InsertAsFullWidth As Variant], [DateLanguage As Variant], [CalendarType As Variant])`  
  Inserts the current date or time, or both, either as text or as a TIME field.
    - `DateTimeFormat As Variant` (optional): The format to be used for displaying the date or time, or both. If this argument is omitted, Microsoft Word uses the short-date style from the Windows Control Panel (Regional Settings icon).
    - `InsertAsField As Variant` (optional): True to insert the specified information as a TIME field. The default value is True.
    - `InsertAsFullWidth As Variant` (optional): True to insert the specified information as double-byte digits. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `DateLanguage As Variant` (optional): Sets the language in which to display the date or time. Can be either of the WdDateLanguage constants. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `CalendarType As Variant` (optional): Sets the calendar type to use when displaying the date or time. Can be either of the WdCalendarTypeBi constants. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
- `Sort([ExcludeHeader As Variant], [FieldNumber As Variant], [SortFieldType As Variant], [SortOrder As Variant], [FieldNumber2 As Variant], [SortFieldType2 As Variant], [SortOrder2 As Variant], [FieldNumber3 As Variant], [SortFieldType3 As Variant], [SortOrder3 As Variant], [SortColumn As Variant], [Separator As Variant], [CaseSensitive As Variant], [BidiSort As Variant], [IgnoreThe As Variant], [IgnoreKashida As Variant], [IgnoreDiacritics As Variant], [IgnoreHe As Variant], [LanguageID As Variant])`  
  Sorts the paragraphs in the specified range.
    - `ExcludeHeader As Variant` (optional): True to exclude the first row or paragraph header from the sort operation. The default value is False.
    - `FieldNumber As Variant` (optional): The fields by which to sort. Microsoft Word sorts by FieldNumber, then by FieldNumber2, and then by FieldNumber3.
    - `SortFieldType As Variant` (optional): The respective sort types for FieldNumber. Can be one of the WdSortFieldType constants. The default value is wdSortFieldAlphanumeric. Some of these constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `SortOrder As Variant` (optional): The sorting order to use when sorting FieldNumber. Can be any WdSortOrder constant.
    - `FieldNumber2 As Variant` (optional): The fields by which to sort.
    - `SortFieldType2 As Variant` (optional): The respective sort types for FieldNumber2. Can be one of the WdSortFieldType constants. The default value is wdSortFieldAlphanumeric. Some of these constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `SortOrder2 As Variant` (optional): The sorting order to use when sorting FieldNumber2. Can be any WdSortOrder constant.
    - `FieldNumber3 As Variant` (optional): The fields by which to sort.
    - `SortFieldType3 As Variant` (optional): Some of these constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed. The default value is wdSortFieldAlphanumeric.
    - `SortOrder3 As Variant` (optional): The sorting order to use when sorting FieldNumber3. Can be any WdSortOrder constant.
    - `SortColumn As Variant` (optional): True to sort only the column specified by the Range object.
    - `Separator As Variant` (optional): The type of field separator. Can be one of the WdSortSeparator constants.
    - `CaseSensitive As Variant` (optional): True to sort with case sensitivity. The default value is False.
    - `BidiSort As Variant` (optional): True to sort based on right-to-left language rules. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `IgnoreThe As Variant` (optional): True to ignore the Arabic character alef lam when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `IgnoreKashida As Variant` (optional): True to ignore kashidas when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `IgnoreDiacritics As Variant` (optional): True to ignore bidirectional control characters when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `IgnoreHe As Variant` (optional): True to ignore the Hebrew character he when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `LanguageID As Variant` (optional): Specifies the sorting language. Can be one of the WdLanguageID constants. Refer to the Object Browser for a list of the WdLanguageID constants.
- `DetectLanguage()`  
  Analyzes the specified text to determine the language that it is written in.
- `ConvertToTable([Separator As Variant], [NumRows As Variant], [NumColumns As Variant], [InitialColumnWidth As Variant], [Format As Variant], [ApplyBorders As Variant], [ApplyShading As Variant], [ApplyFont As Variant], [ApplyColor As Variant], [ApplyHeadingRows As Variant], [ApplyLastRow As Variant], [ApplyFirstColumn As Variant], [ApplyLastColumn As Variant], [AutoFit As Variant], [AutoFitBehavior As Variant], [DefaultTableBehavior As Variant]) As Table`  
  Converts text within a range to a table. Returns the table as a Table object.
    - `Separator As Variant` (optional): Specifies the character used to separate text into cells. Can be a character or one of the following WdTableFieldSeparator constant. If this argument is omitted, the value of the DefaultTableSeparator property is used.
    - `NumRows As Variant` (optional): The number of rows in the table. If this argument is omitted, Microsoft Word sets the number of rows, based on the contents of the range.
    - `NumColumns As Variant` (optional): The number of columns in the table. If this argument is omitted, Word sets the number of columns, based on the contents of the range.
    - `InitialColumnWidth As Variant` (optional): The initial width of each column, in points. If this argument is omitted, Word calculates and adjusts the column width so that the table stretches from margin to margin.
    - `Format As Variant` (optional): Specifies one of the predefined formats listed in the Table AutoFormat dialog box. Can be one of the WdTableFormat constants.
    - `ApplyBorders As Variant` (optional): True to apply the border properties of the specified format.
    - `ApplyShading As Variant` (optional): True to apply the shading properties of the specified format.
    - `ApplyFont As Variant` (optional): True to apply the font properties of the specified format.
    - `ApplyColor As Variant` (optional): True to apply the color properties of the specified format.
    - `ApplyHeadingRows As Variant` (optional): True to apply the heading-row properties of the specified format.
    - `ApplyLastRow As Variant` (optional): True to apply the last-row properties of the specified format.
    - `ApplyFirstColumn As Variant` (optional): True to apply the first-column properties of the specified format.
    - `ApplyLastColumn As Variant` (optional): True to apply the last-column properties of the specified format.
    - `AutoFit As Variant` (optional): True to decrease the width of the table columns as much as possible without changing the way text wraps in the cells.
    - `AutoFitBehavior As Variant` (optional): Sets the AutoFit rules for how Word sizes a table. Can be one of the following WdAutoFitBehavior constant. If DefaultTableBehavior is wdWord8TableBehavior, this argument is ignored.
    - `DefaultTableBehavior As Variant` (optional): Sets a value that specifies whether Microsoft Word automatically resizes cells in a table to fit the contents (AutoFit). Can be one of the WdDefaultTableBehavior constant.
- `TCSCConverter([WdTCSCConverterDirection As WdTCSCConverterDirection], [CommonTerms As Boolean], [UseVariants As Boolean])`  
  Converts the specified range from Traditional Chinese to Simplified Chinese or vice versa.
    - `WdTCSCConverterDirection As WdTCSCConverterDirection` (optional): Specifies the direction in which text is converted. If omitted, the default value is wdTCSCConverterDirectionAuto, which converts in the appropriate direction based on the detected language of the specified range.
    - `UseVariants As Boolean` (optional): True if Word uses Taiwan, Hong Kong SAR, and Macao SAR character variants. Can only be used if translating from Simplified Chinese to Traditional Chinese.
- `PasteAndFormat(Type As WdRecoveryType)`  
  Pastes the selected table cells and formats them as specified.
    - `Type As WdRecoveryType` (required): The type of formatting to use when pasting the selected table cells.
- `PasteExcelTable(LinkedToExcel As Boolean, WordFormatting As Boolean, RTF As Boolean)`  
  Pastes and formats a Microsoft Excel table.
    - `LinkedToExcel As Boolean` (required): True links the pasted table to the original Excel file so that changes made to the Excel file are reflected in Microsoft Word.
    - `WordFormatting As Boolean` (required): True formats the table using the formatting in the Word document. False formats the table according to the original Excel file.
    - `RTF As Boolean` (required): True pastes the Excel table using Rich Text Format (RTF). False pastes the Excel table as HTML.
- `PasteAppendTable()`  
  Merges pasted cells into an existing table by inserting the pasted rows between the selected rows. No cells are overwritten.
- `GoToEditableRange([EditorID As Variant]) As Range`  
  Returns a Range object that represents an area of a document that can be modified by the specified user or group of users.
- `InsertXML(XML As String, [Transform As Variant])`  
  Inserts the specified XML into the document at the specified range, replacing any text contained within the range.
    - `XML As String` (required): Specifies the XML to insert. This can be any valid custom XML.
    - `Transform As Variant` (optional): Specifies the XML Transformation (XSLT) used to transform the XML. If omitted, the XML is inserted as custom XML without applying a transform.
- `InsertCaption(Label As Variant, [Title As Variant], [TitleAutoText As Variant], [Position As Variant], [ExcludeLabel As Variant])`  
  Inserts a caption immediately preceding or following the specified range.
    - `Label As Variant` (required): The caption label to be inserted. Can be a String or one of the WdCaptionLabelID constants. If the label has not yet been defined, an error occurs. Use the Add method with the CaptionLabels object to define new caption labels.
    - `Title As Variant` (optional): The string to be inserted immediately following the label in the caption (ignored if TitleAutoText is specified).
    - `TitleAutoText As Variant` (optional): The AutoText entry whose contents you want to insert immediately following the label in the caption (overrides any text specified by Title).
    - `Position As Variant` (optional): Specifies whether the caption will be inserted above or below the range. Can be either one of the WdCaptionPosition constants.
    - `ExcludeLabel As Variant` (optional): True does not include the text label, as defined in the Label parameter. False includes the specified label.
- `InsertCrossReference(ReferenceType As Variant, ReferenceKind As WdReferenceKind, ReferenceItem As Variant, [InsertAsHyperlink As Variant], [IncludePosition As Variant], [SeparateNumbers As Variant], [SeparatorString As Variant])`  
  Inserts a cross-reference to a heading, bookmark, footnote, or endnote, or to an item for which a caption label is defined (for example, an equation, figure, or table).
    - `ReferenceType As Variant` (required): The type of item for which a cross-reference is to be inserted. Can be any WdReferenceType or WdCaptionLabelID constant or a user defined caption label.
    - `ReferenceKind As WdReferenceKind` (required): The information to be included in the cross-reference.
    - `ReferenceItem As Variant` (required): If ReferenceType is wdRefTypeBookmark, this argument specifies a bookmark name. For all other ReferenceType values, this argument specifies the item number or name in the Reference type option in the Cross-reference dialog box. Use the GetCrossReferenceItems method to return a list of item names that can be used with this argument.
    - `InsertAsHyperlink As Variant` (optional): True to insert the cross-reference as a hyperlink to the referenced item.
    - `IncludePosition As Variant` (optional): True to insert "above" or "below," depending on the location of the reference item in relation to the cross-reference.
    - `SeparateNumbers As Variant` (optional): True to use a separator to separate the numbers from the associated text. (Use only if the ReferenceType parameter is set to wdRefTypeNumberedItem and the ReferenceKind parameter is set to wdNumberFullContext.)
    - `SeparatorString As Variant` (optional): Specifies the string to use as a separator if the SeparateNumbers parameter is set to True.
- `ExportFragment(FileName As String, Format As WdSaveFormat)`  
  Exports the selected range into a document for use as a document fragment.
    - `FileName As String` (required): Specifies the path and file name of the file in which to save the document fragment.
    - `Format As WdSaveFormat` (required): Specifies the file format of the document fragment file.
- `SetListLevel(Level As Integer)`  
  Sets the list level for one or more items in a numbered list.
    - `Level As Integer` (required): A number that indicates the new list level.
- `InsertAlignmentTab(Alignment As Long, [RelativeTo As Long])`  
  Inserts an absolute tab that is always positioned in the same spot, relative to either the margins or indents.
    - `Alignment As Long` (required): Indicates the type of alignment&mdash;left, center, or right&mdash;for the tab stop. Can be one of the WdAlignmentTabAlignment constants.
    - `RelativeTo As Long` (optional): Indicates whether the tab stop is relative to the margins or to the paragraph indents. Can be one of the WdAlignmentTabRelative constants.
- `ImportFragment(FileName As String, [MatchDestination As Boolean])`  
  Imports a document fragment into the document at the specified range.
    - `FileName As String` (required): Specifies the path and file name where the document fragment is stored.
    - `MatchDestination As Boolean` (optional): Specifies whether to match the destination formatting. If False, the imported document fragment retains the formatting in the original document. Default value is False.
- `ExportAsFixedFormat(OutputFileName As String, ExportFormat As WdExportFormat, [OpenAfterExport As Boolean], [OptimizeFor As WdExportOptimizeFor], [ExportCurrentPage As Boolean], [Item As WdExportItem], [IncludeDocProps As Boolean], [KeepIRM As Boolean], [CreateBookmarks As WdExportCreateBookmarks], [DocStructureTags As Boolean], [BitmapMissingFonts As Boolean], [UseISO19005_1 As Boolean], [FixedFormatExtClassPtr As Variant])`  
  Saves a portion of a document as PDF or XPS format.
    - `OutputFileName As String` (required): The path and file name of the new PDF or XPS file.
    - `ExportFormat As WdExportFormat` (required): Specifies either PDF or XPS format.
    - `OpenAfterExport As Boolean` (optional): Opens the new file after exporting the contents.
    - `OptimizeFor As WdExportOptimizeFor` (optional): Specifies whether to optimize for screen or print.
    - `ExportCurrentPage As Boolean` (optional): Specifies whether to export the current page. True exports the entire page. False exports only the current selection.
    - `Item As WdExportItem` (optional): Specifies whether the export process includes text only or includes text with markup.
    - `IncludeDocProps As Boolean` (optional): Specifies whether to include document properties in the newly exported file.
    - `KeepIRM As Boolean` (optional): Specifies whether to copy IRM permissions to an XPS document if the source document has IRM protections. Default value is True.
    - `CreateBookmarks As WdExportCreateBookmarks` (optional): Specifies whether to export bookmarks and the type of bookmarks to export.
    - `DocStructureTags As Boolean` (optional): Specifies whether to include extra data to help screen readers, for example information about the flow and logical organization of the content. Default value is True.
    - `BitmapMissingFonts As Boolean` (optional): Specifies whether to include a bitmap of the text. Set this parameter to True when font licenses don't permit a font to be embedded in the PDF file. If False, the font is referenced, and the viewer's computer substitutes an appropriate font if the authored one is not available. Default value is True.
    - `FixedFormatExtClassPtr As Variant` (optional): Specifies a pointer to an add-in that allows calls to an alternate implementation of code. The alternate implementation of code interprets the EMF and EMF+ page descriptions that are generated by the applications to make their own PDF or XPS. For more information, see Extend the fixed-format export feature in Word Automation Services.
- `SortByHeadings([SortFieldType As Variant], [SortOrder As Variant], [CaseSensitive As Variant], [BidiSort As Variant], [IgnoreThe As Variant], [IgnoreKashida As Variant], [IgnoreDiacritics As Variant], [IgnoreHe As Variant], [LanguageID As Variant])`  
  Sorts the headings in the specified range.
    - `SortFieldType As Variant` (optional): The sort field type to use. Can be one of the WdSortFieldType constants. The default value is wdSortFieldAlphanumeric. Depending on the language support (U.S. English, for example) that you have selected or installed, some of these constants may not be available to you.
    - `SortOrder As Variant` (optional): The sorting order to use. Can be one of the WdSortOrder constants.
    - `CaseSensitive As Variant` (optional): True to sort with case sensitivity. The default value is False.
    - `BidiSort As Variant` (optional): True to sort based on right-to-left language rules. Depending on the language support (U.S. English, for example) that you have selected or installed, this parameter may not be available to you.
    - `IgnoreThe As Variant` (optional): True to ignore the Arabic character alef lam when sorting right-to-left language text. Depending on the language support (U.S. English, for example) that you have selected or installed, this parameter may not be available to you.
    - `IgnoreKashida As Variant` (optional): True to ignore kashidas when sorting right-to-left language text. Depending on the language support (U.S. English, for example) that you have selected or installed, this parameter may not be available to you.
    - `IgnoreDiacritics As Variant` (optional): True to ignore bidirectional control characters when sorting right-to-left language text. Depending on the language support (U.S. English, for example) that you have selected or installed, this parameter may not be available to you.
    - `IgnoreHe As Variant` (optional): True to ignore the Hebrew character he when sorting right-to-left language text. Depending on the language support (U.S. English, for example) that you have selected or installed, this parameter may not be available to you.
    - `LanguageID As Variant` (optional): Specifies the sorting language. Can be one of the WdLanguageID constants.
- `ExportAsFixedFormat2(OutputFileName As String, ExportFormat As WdExportFormat, [OpenAfterExport As Boolean], [OptimizeFor As WdExportOptimizeFor], [ExportCurrentPage As Boolean], [Item As WdExportItem], [IncludeDocProps As Boolean], [KeepIRM As Boolean], [CreateBookmarks As WdExportCreateBookmarks], [DocStructureTags As Boolean], [BitmapMissingFonts As Boolean], [UseISO19005_1 As Boolean], [OptimizeForImageQuality As Boolean], [FixedFormatExtClassPtr As Variant])`  
  Saves a portion of a document as PDF or XPS format.
    - `OutputFileName As String` (required): The path and file name of the new PDF or XPS file.
    - `ExportFormat As WdExportFormat` (required): Specifies either PDF or XPS format.
    - `OpenAfterExport As Boolean` (optional): Opens the new file after exporting the contents.
    - `OptimizeFor As WdExportOptimizeFor` (optional): Specifies whether to optimize for screen or print.
    - `ExportCurrentPage As Boolean` (optional): Specifies whether to export the current page. True exports the entire page. False exports only the current selection.
    - `Item As WdExportItem` (optional): Specifies whether the export process includes text only or includes text with markup.
    - `IncludeDocProps As Boolean` (optional): Specifies whether to include document properties in the newly exported file.
    - `KeepIRM As Boolean` (optional): Specifies whether to copy IRM permissions to an XPS document if the source document has IRM protections. Default value is True.
    - `CreateBookmarks As WdExportCreateBookmarks` (optional): Specifies whether to export bookmarks and the type of bookmarks to export.
    - `DocStructureTags As Boolean` (optional): Specifies whether to include extra data to help screen readers, for example information about the flow and logical organization of the content. Default value is True.
    - `BitmapMissingFonts As Boolean` (optional): Specifies whether to include a bitmap of the text. Set this parameter to True when font licenses don't permit a font to be embedded in the PDF file. If False, the font is referenced, and the viewer's computer substitutes an appropriate font if the authored one is not available. Default value is True.
    - `UseISO19005_1 As Boolean` (optional): Specifies whether to limit PDF usage to the PDF subset standardized as ISO 19005-1. If True, the resulting files are more reliably self-contained but may be larger or show more visual artifacts due to the restrictions of the format. Default value is False.
    - `OptimizeForImageQuality As Boolean` (optional): Specifies whether to downsample images or keep their original quality. If True, the resulting files will have better image quality but may be larger. Default value is False.
    - `FixedFormatExtClassPtr As Variant` (optional): Specifies a pointer to an add-in that allows calls to an alternate implementation of code. The alternate implementation of code interprets the EMF and EMF+ page descriptions that are generated by the applications to make their own PDF or XPS. For more information, see Extend the fixed-format export feature in Word Automation Services.
- `ExportAsFixedFormat3(OutputFileName As String, ExportFormat As WdExportFormat, [OpenAfterExport As Boolean], [OptimizeFor As WdExportOptimizeFor], [ExportCurrentPage As Boolean], [Item As WdExportItem], [IncludeDocProps As Boolean], [KeepIRM As Boolean], [CreateBookmarks As WdExportCreateBookmarks], [DocStructureTags As Boolean], [BitmapMissingFonts As Boolean], [UseISO19005_1 As Boolean], [OptimizeForImageQuality As Boolean], [ImproveExportTagging As Boolean], [FixedFormatExtClassPtr As Variant])`  
  Saves a portion of a document as PDF or XPS format.
    - `OutputFileName As String` (required): The path and file name of the new PDF or XPS file.
    - `ExportFormat As WdExportFormat` (required): Specifies either PDF or XPS format.
    - `OpenAfterExport As Boolean` (optional): Opens the new file after exporting the contents.
    - `OptimizeFor As WdExportOptimizeFor` (optional): Specifies whether to optimize for screen or print.
    - `ExportCurrentPage As Boolean` (optional): Specifies whether to export the current page. True exports the entire page. False exports only the current selection.
    - `Item As WdExportItem` (optional): Specifies whether the export process includes text only or includes text with markup.
    - `IncludeDocProps As Boolean` (optional): Specifies whether to include document properties in the newly exported file.
    - `KeepIRM As Boolean` (optional): Specifies whether to copy IRM permissions to an XPS document if the source document has IRM protections. Default value is True.
    - `CreateBookmarks As WdExportCreateBookmarks` (optional): Specifies whether to export bookmarks and the type of bookmarks to export.
    - `DocStructureTags As Boolean` (optional): Specifies whether to include extra data to help screen readers, for example information about the flow and logical organization of the content. Default value is True.
    - `BitmapMissingFonts As Boolean` (optional): Specifies whether to include a bitmap of the text. Set this parameter to True when font licenses don't permit a font to be embedded in the PDF file. If False, the font is referenced, and the viewer's computer substitutes an appropriate font if the authored one is not available. Default value is True.
    - `OptimizeForImageQuality As Boolean` (optional): Specifies whether to downsample images or keep their original quality. If True, the resulting files will have better image quality but may be larger. Default value is False.
    - `ImproveExportTagging As Boolean` (optional): Specifies whether to enable improved accessbility tagging. For more information, see the Remarks section. Default value is False.
    - `FixedFormatExtClassPtr As Variant` (optional): Specifies a pointer to an add-in that allows calls to an alternate implementation of code. The alternate implementation of code interprets the EMF and EMF+ page descriptions that are generated by the applications to make their own PDF or XPS.
