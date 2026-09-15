# Selection

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020975-0000-0000-C000-000000000046}  

Represents the current selection in a window or pane. A selection represents either a selected (or highlighted) area in the document, or it represents the insertion point if nothing in the document is selected. There can be only one Selection object per document window pane, and only one Selection object in the entire application can be active.

**Remarks:** Use the Selection property to return the Selection object. If no object qualifier is used with the Selection property, Microsoft Word returns the selection from the active pane of the active document window. The following example copies the current selection from the active document. The following example deletes the selection from the third document in the Documents collection. The document does not have to be active to access its current selection. The following example copies the selection from the first pane of the active document and pastes it into the second pane. The Text property is the default property of the Selection object. Use this property to set or return the text in the current selection. The following example assigns the text in the current selection to the variable strTemp, removing the last character if it is a paragraph mark. The Selection object has various methods and properties with which you can collapse, expand, or otherwise change the current selection. The following example moves the insertion point to the end of the document and selects the last three lines.

## Properties (67)

- `Text As String  (read/write)`  
  Returns or sets the text in the specified selection. Read/write String.
- `FormattedText As Range  (read/write)`  
  Returns or sets a Range object that includes the formatted text in the specified range or selection. Read/write.
- `Start As Long  (read/write)`  
  Returns or sets the starting character position of a selection. Read/write Long.
- `End As Long  (read/write)`  
  Returns or sets the ending character position of a selection. Read/write Long.
- `Font As Font  (read/write)`  
  Returns or sets a Font object that represents the character formatting of the specified object. Read/write.
- `Type As WdSelectionType  (read-only)`  
  Returns the selection type. Read-only WdSelectionType.
- `StoryType As WdStoryType  (read-only)`  
  Returns the story type for the specified selection. Read-only WdStoryType.
- `Style As Variant  (read/write)`  
  Returns or sets the style for the specified object. To set this property, specify the local name of the style, an integer, a WdBuiltinStyle constant, or an object that represents the style. For a list of valid constants, consult the Microsoft Visual Basic Object Browser. Read/write Variant.
- `Tables As Tables  (read-only)`  
  Returns a Tables collection that represents all the tables in the specified selection. Read-only.
- `Words As Words  (read-only)`  
  Returns a Words collection that represents all the words in a selection. Read-only.
- `Sentences As Sentences  (read-only)`  
  Returns a Sentences collection that represents all the sentences in the selection. Read-only.
- `Characters As Characters  (read-only)`  
  Returns a Characters collection that represents the characters in a document, range, or selection. Read-only.
- `Footnotes As Footnotes  (read-only)`  
  Returns a Footnotes collection that represents all the footnotes in a range, selection, or document. Read-only.
- `Endnotes As Endnotes  (read-only)`  
  Returns an Endnotes collection that represents all the endnotes contained within a selection. Read-only.
- `Comments As Comments  (read-only)`  
  Returns a Comments collection that represents all the comments in the specified. Read-only.
- `Cells As Cells  (read-only)`  
  Returns a Cells collection that represents the table cells in a selection. Read-only.
- `Sections As Sections  (read-only)`  
  Returns a Sections collection that represents the sections in the specified selection. Read-only.
- `Paragraphs As Paragraphs  (read-only)`  
  Returns a Paragraphs collection that represents all the paragraphs in the specified selection. Read-only.
- `Borders As Borders  (read/write)`  
  Returns a Borders collection that represents all the borders for the specified object.
- `Shading As Shading  (read-only)`  
  Returns a Shading object that refers to the shading formatting for the specified selection.
- `Fields As Fields  (read-only)`  
  Returns a read-only Fields collection that represents all the fields in the selection.
- `FormFields As FormFields  (read-only)`  
  Returns a FormFields collection that represents all the form fields in the selection. Read-only.
- `Frames As Frames  (read-only)`  
  Returns a Frames collection that represents all the frames in a selection. Read-only.
- `ParagraphFormat As ParagraphFormat  (read/write)`  
  Returns or sets a ParagraphFormat object that represents the paragraph settings for the specified selection. Read/write.
- `PageSetup As PageSetup  (read/write)`  
  Returns a PageSetup object that's associated with the specified selection.
- `Bookmarks As Bookmarks  (read-only)`  
  Returns a Bookmarks collection that represents all the bookmarks in a document, range, or selection. Read-only.
- `StoryLength As Long  (read-only)`  
  Returns the number of characters in the story that contains the specified selection. Read-only Long.
- `LanguageID As WdLanguageID  (read/write)`  
  Returns or sets the language for the specified object. Read/write .
- `LanguageIDFarEast As WdLanguageID  (read/write)`  
  Returns or sets an East Asian language for the specified object. Read/write WdLanguageID.
- `LanguageIDOther As WdLanguageID  (read/write)`  
  Returns or sets the language for the specified object. Read/write WdLanguageID.
- `Hyperlinks As Hyperlinks  (read-only)`  
  Returns a Hyperlinks collection that represents all the hyperlinks in the specified selection. Read-only.
- `Columns As Columns  (read-only)`  
  Returns a Columns collection that represents all the table columns in a selection. Read-only.
- `Rows As Rows  (read-only)`  
  Returns a Rows collection that represents all the table rows in a range, selection, or table. Read-only.
- `HeaderFooter As HeaderFooter  (read-only)`  
  Returns a HeaderFooter object for the specified selection. Read-only.
- `IsEndOfRowMark As Boolean  (read-only)`  
  True if the specified selection or range is collapsed and is located at the end-of-row mark in a table. Read-only Boolean.
- `BookmarkID As Long  (read-only)`  
  Returns the number of the bookmark that encloses the beginning of the specified selection. Read-only Long.
- `PreviousBookmarkID As Long  (read-only)`  
  Returns the number of the last bookmark that starts before or at the same place as the specified selection or range; returns 0 (zero) if there is no corresponding bookmark. Read-only Long.
- `Find As Find  (read-only)`  
  Returns a Find object that contains the criteria for a find operation. Read-only.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that's contained in the specified object.
- `Information As Variant  (read-only)`  
  Returns information about the specified selection. Read-only Variant.
- `Flags As WdSelectionFlags  (read/write)`  
  Returns or sets properties of the selection. Read/write WdSelectionFlags.
- `Active As Boolean  (read-only)`  
  True if the selection in the specified window or pane is active. Read-only Boolean.
- `StartIsActive As Boolean  (read/write)`  
  True if the beginning of the selection is active. Read/write Boolean.
- `IPAtEndOfLine As Boolean  (read-only)`  
  True if the insertion point is at the end of a line that wraps to the next line. Read-only Boolean.
- `ExtendMode As Boolean  (read/write)`  
  True if Extend mode is active. Read/write Boolean.
- `ColumnSelectMode As Boolean  (read/write)`  
  True if column selection mode is active. Read/write Boolean.
- `Orientation As WdTextOrientation  (read/write)`  
  Returns or sets the orientation of text in a selection when the Text Direction feature is enabled. Read/write WdTextOrientation.
- `InlineShapes As InlineShapes  (read-only)`  
  Returns an InlineShapes collection that represents all the InlineShape objects in a selection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Selection object.
- `Document As Document  (read-only)`  
  Returns a Document object associated with the specified selection. Read-only.
- `ShapeRange As ShapeRange  (read-only)`  
  Returns a ShapeRange collection that represents all the Shape objects in the selection. Read-only.
- `NoProofing As Long  (read/write)`  
  True if the spelling and grammar checker ignores the specified text. Returns wdUndefined if the NoProofing property is set to True for only some of the specified text. Read/write Long.
- `TopLevelTables As Tables  (read-only)`  
  Returns a Tables collection that represents the tables at the outermost nesting level in the current selection. Read-only.
- `LanguageDetected As Boolean  (read/write)`  
  Returns or sets a Boolean that specifies whether Microsoft Word has detected the language of the selected text.
- `FitTextWidth As Single  (read/write)`  
  Returns or sets the width (in the current measurement units) in which Microsoft Word fits the text in the current selection. Read/write Single.
- `HTMLDivisions As HTMLDivisions  (read-only)`  
  Returns an HTMLDivisions object that represents an HTML division in a web document.
- `ChildShapeRange As ShapeRange  (read-only)`  
  Returns a ShapeRange collection representing the child shapes contained within a selection.
- `HasChildShapeRange As Boolean  (read-only)`  
  True if the selection contains child shapes. Read-only Boolean.
- `FootnoteOptions As FootnoteOptions  (read-only)`  
  Returns FootnoteOptions object that represents the footnotes in a selection.
- `EndnoteOptions As EndnoteOptions  (read-only)`  
  Returns an EndnoteOptions object that represents the endnotes in a selection.
- `Editors As Editors  (read-only)`  
  Returns an Editors object that represents all the users authorized to modify a selection within a document.
- `XML As String  (read-only)`  
  Returns a String that represents the XML text in the specified object. .
- `EnhMetaFileBits As Variant  (read-only)`  
  Returns a Variant that represents a picture representation of how a selection or range of text appears.
- `OMaths As OMaths  (read-only)`  
  Returns an OMaths collection that represents the OMath objects within the current selection. Read-only.
- `WordOpenXML As String  (read-only)`  
  Returns a String that represents the XML contained within the selection in the Microsoft Word Open XML format. Read-only.

## Methods (115)

- `Select()`  
  Selects the specified text.
- `SetRange(Start As Long, End As Long)`  
  Sets the starting and ending character positions for the selection.
    - `Start As Long` (required): The starting character position of the selection.
    - `End As Long` (required): The ending character position of the selection.
- `Collapse([Direction As Variant])`  
  Collapses a selection to the starting or ending position. After a selection is collapsed, the starting and ending points are equal.
    - `Direction As Variant` (optional): The direction in which to collapse the range or selection. Can be either of the following WdCollapseDirection constants: wdCollapseEnd or wdCollapseStart. The default value is wdCollapseStart.
- `InsertBefore(Text As String)`  
  Inserts the specified text before the specified selection. .
    - `Text As String` (required): The text to be inserted.
- `InsertAfter(Text As String)`  
  Inserts the specified text at the end of a range or selection.
    - `Text As String` (required): The text to be inserted.
- `Next([Unit As Variant], [Count As Variant]) As Range`  
  Returns a Range object that represents the next unit relative to the specified selection.
    - `Unit As Variant` (optional): The type of units by which to count. Can be any WdUnits constant.
    - `Count As Variant` (optional): The number of units by which you want to move ahead. The default value is one.
- `Previous([Unit As Variant], [Count As Variant]) As Range`  
  Moves the selected text by the specified number of units, and returns a Range object relative to the collapsed selection.
    - `Unit As Variant` (optional): Specifies the type of unit by which to move the selection. Can be one of the WdUnits constants.
    - `Count As Variant` (optional): The number of units by which you want to move. The default value is 1.
- `StartOf([Unit As Variant], [Extend As Variant]) As Long`  
  Moves or extends the start position of the specified range or selection to the beginning of the nearest specified text unit. This method returns a Long that indicates the number of characters by which the range or selection was moved or extended. The method returns a negative number if the movement is backward through the document.
    - `Unit As Variant` (optional): The unit by which the start position of the specified range or selection is to be moved. If a value is omitted, the default value is wdWord.
    - `Extend As Variant` (optional): If you use wdMove, both ends of the range or selection are moved to the beginning of the specified unit. If you use wdExtend, the beginning of the range or selection is extended to the beginning of the specified unit. The default value is wdMove.
- `EndOf([Unit As Variant], [Extend As Variant]) As Long`  
  Moves or extends the ending character position of a range or selection to the end of the nearest specified text unit.
    - `Unit As Variant` (optional): The unit by which to move the ending character position. WdUnits.
    - `Extend As Variant` (optional): Can be either of the WdMovementType constants. If wdMove, both ends of the range or selection object are moved to the end of the specified unit. If wdExtend is used, the end of the range or selection is extended to the end of the specified unit. The default value is wdMove.
- `Move([Unit As Variant], [Count As Variant]) As Long`  
  Collapses the specified selection to its start or end position and then moves the collapsed object by the specified number of units. This method returns a Long value that represents the number of units by which the selection was moved, or it returns 0 (zero) if the move was unsuccessful.
    - `Unit As Variant` (optional): The unit by which to move the ending character position.
    - `Count As Variant` (optional): The number of units by which the specified range or selection is to be moved. If Count is a positive number, the object is collapsed to its end position and moved backward in the document by the specified number of units. If Count is a negative number, the object is collapsed to its start position and moved forward by the specified number of units. The default value is 1. You can also control the collapse direction by using the Collapse method before using the Move method. If the range or selection is in the middle of a unit or isn't collapsed, moving it to the beginning or end of the unit counts as moving it one full unit.
- `MoveStart([Unit As Variant], [Count As Variant]) As Long`  
  Moves the start position of the specified selection.
    - `Unit As Variant` (optional): The unit by which start position of the specified selection is to be moved. Can be one of the WdUnits constants. The default value is wdCharacter.
    - `Count As Variant` (optional): The maximum number of units by which the specified selection is to be moved. If Count is a positive number, the start position of the selection is moved forward in the document. If it is a negative number, the start position is moved backward. If the start position is moved forward to a position beyond the end position, the selection is collapsed and both the start and end positions are moved together. The default value is 1.
- `MoveEnd([Unit As Variant], [Count As Variant]) As Long`  
  Moves the ending character position of a range or selection.
    - `Unit As Variant` (optional): The unit by which to move the ending character position. The default value is wdCharacter.
    - `Count As Variant` (optional): The number of units to move. If this number is positive, the ending character position is moved forward in the document. If this number is negative, the end is moved backward. If the ending position overtakes the starting position, the range collapses and both character positions move together. The default value is 1.
- `MoveWhile(Cset As Variant, [Count As Variant]) As Long`  
  Moves the specified selection while any of the specified characters are found in the document.
    - `Cset As Variant` (required): One or more characters. This argument is case-sensitive.
    - `Count As Variant` (optional): The maximum number of characters by which the specified selection is to be moved. Can be a number or either the wdForward or wdBackward constant. If Count is a positive number, the specified selection is moved forward in the document, beginning at the end position. If it is a negative number, the selection is moved backward, beginning at the start position. The default value is wdForward.
- `MoveStartWhile(Cset As Variant, [Count As Variant]) As Long`  
  Moves the start position of the specified selection while any of the specified characters are found in the document.
    - `Cset As Variant` (required): One or more characters. This argument is case-sensitive.
    - `Count As Variant` (optional): The maximum number of characters by which the specified selection is to be moved. Can be a number or either the wdForward or wdBackward constant. If Count is a positive number, the selection is moved forward in the document. If it is a negative number, the selection is moved backward. The default value is wdForward.
- `MoveEndWhile(Cset As Variant, [Count As Variant]) As Long`  
  Moves the ending character position of a selection while any of the specified characters are found in the document.
    - `Cset As Variant` (required): One or more characters. This argument is case-sensitive.
    - `Count As Variant` (optional): The maximum number of characters by which the selection is to be moved. Can be a number or either wdForward or wdBackward. If Count is a positive number, the selection is moved forward in the document. If it is a negative number, the selection is moved backward. The default value is wdForward.
- `MoveUntil(Cset As Variant, [Count As Variant]) As Long`  
  Moves the specified selection until one of the specified characters is found in the document.
    - `Cset As Variant` (required): One or more characters. If any character in Cset is found before the Count value expires, the specified selection is positioned as an insertion point immediately before that character. This argument is case-sensitive.
    - `Count As Variant` (optional): The maximum number of characters by which the specified selection is to be moved. Can be a number or either the wdForward or wdBackward constant. If Count is a positive number, the selection is moved forward in the document, beginning at the end position. If it is a negative number, the selection is moved backward, beginning at the start position. The default value is wdForward.
- `MoveStartUntil(Cset As Variant, [Count As Variant]) As Long`  
  Moves the start position of the specified selection until one of the specified characters is found in the document. If the movement is backward through the document, the selection is expanded.
    - `Cset As Variant` (required): One or more characters. This argument is case-sensitive.
    - `Count As Variant` (optional): The maximum number of characters by which the specified selection is to be moved. Can be a number or either the wdForward or wdBackward constant. If Count is a positive number, the selection is moved forward in the document. If it is a negative number, the selection is moved backward. The default value is wdForward.
- `MoveEndUntil(Cset As Variant, [Count As Variant]) As Long`  
  Moves the end position of the specified selection until any of the specified characters are found in the document.
    - `Cset As Variant` (required): One or more characters. This argument is case-sensitive.
    - `Count As Variant` (optional): The maximum number of characters by which the specified selection is to be moved. Can be a number or either wdForward or wdBackward. If Count is a positive number, the selection is moved forward in the document. If it is a negative number, the selection is moved backward. The default value is wdForward.
- `Cut()`  
  Removes the specified object from the document and moves it to the Clipboard.
- `Copy()`  
  Copies the specified selection to the Clipboard.
- `Paste()`  
  Inserts the contents of the Clipboard at the specified selection.
- `InsertBreak([Type As Variant])`  
  Inserts a page, column, or section break.
    - `Type As Variant` (optional): the type of break to insert. The default value is wdPageBreak. Some of the WdBreakType constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
- `InsertFile(FileName As String, [Range As Variant], [ConfirmConversions As Variant], [Link As Variant], [Attachment As Variant])`  
  Inserts all or part of the specified file.
    - `FileName As String` (required): The path and file name of the file to be inserted. If you don't specify a path, Word assumes the file is in the current folder.
    - `Range As Variant` (optional): If the specified file is a Word document, this parameter refers to a bookmark. If the file is another type (for example, a Microsoft Excel worksheet), this parameter refers to a named range or a cell range (for example, R1C1:R3C4).
    - `ConfirmConversions As Variant` (optional): True to have Word prompt you to confirm conversion when inserting files in formats other than the Word Document format.
    - `Link As Variant` (optional): True to insert the file by using an INCLUDETEXT field.
    - `Attachment As Variant` (optional): True to insert the file as an attachment to an email message.
- `InStory(Range As Range) As Boolean`  
  True if the selection to which this method is applied is in the same story as the range specified by the Range argument.
    - `Range As Range` (required): The range whose story is compared with the story of the current selection.
- `InRange(Range As Range) As Boolean`  
  True if the selection to which the method is applied is contained within the range specified by the Range argument.
    - `Range As Range` (required): The Range to which you want to compare the selection.
- `Delete([Unit As Variant], [Count As Variant]) As Long`  
  Deletes the specified number of characters or words.
    - `Unit As Variant` (optional): The unit by which the collapsed selection is to be deleted. Can be one of the WdUnits constants.
    - `Count As Variant` (optional): The number of units to be deleted. To delete units after the selection, collapse the selection and use a positive number. To delete units before the selection, collapse the selection and use a negative number.
- `Expand([Unit As Variant]) As Long`  
  Expands the specified range or selection. Returns the number of characters added to the range or selection. Long.
    - `Unit As Variant` (optional): A WdUnits constant that represents the unit by which to expand the range. The default value is wdWord.
- `InsertParagraph()`  
  Replaces the specified selection with a new paragraph.
- `InsertParagraphAfter()`  
  Inserts a paragraph mark after a selection.
- `InsertSymbol(CharacterNumber As Long, [Font As Variant], [Unicode As Variant], [Bias As Variant])`  
  Inserts a symbol in place of the specified selection.
    - `CharacterNumber As Long` (required): The character number for the specified symbol. This value will always be the sum of 31 and the number that corresponds to the position of the symbol in the table of symbols (counting from left to right). For example, to specify a delta character at position 37 in the table of symbols in the Symbol font, set CharacterNumber to 68.
    - `Font As Variant` (optional): The name of the font that contains the symbol.
    - `Unicode As Variant` (optional): True to insert the unicode character specified by CharacterNumber; False to insert the ANSI character specified by CharacterNumber. The default value is False.
    - `Bias As Variant` (optional): Sets the font bias for symbols. This argument is useful for setting the correct font bias for East Asian characters. Can be one of the WdFontBias constants. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
- `CopyAsPicture()`  
  The CopyAsPicture method works the same way as the Copy method.
- `SortAscending()`  
  Sorts paragraphs or table rows in ascending alphanumeric order.
- `SortDescending()`  
  Sorts paragraphs or table rows within the selection in descending alphanumeric order.
- `IsEqual(Range As Range) As Boolean`  
  True if the selection to which this method is applied is equal to the range specified by the Range argument.
    - `Range As Range` (required): The range to compare with the current selection.
- `Calculate() As Single`  
  Calculates a mathematical expression within a selection. Returns the result as a Single.
- `GoTo([What As Variant], [Which As Variant], [Count As Variant], [Name As Variant]) As Range`  
  Moves the insertion point to the character position immediately preceding the specified item, and returns a Range object (except for the wdGoToGrammaticalError, wdGoToProofreadingError, or wdGoToSpellingError constant).
    - `What As Variant` (optional): The kind of item to which the range or selection is moved. Can be one of the WdGoToItem constants.
    - `Which As Variant` (optional): The item to which the range or selection is moved. Can be one of the WdGoToDirection constants.
    - `Count As Variant` (optional): The number of the item in the document. The default value is 1. Only positive values are valid. To specify an item that precedes the range or selection, use wdGoToPrevious as the Which argument and specify a Count value.
    - `Name As Variant` (optional): If the What argument is wdGoToBookmark, wdGoToComment, wdGoToField, or wdGoToObject, this argument specifies a name.
- `GoToNext(What As WdGoToItem) As Range`  
  Returns a Range object that refers to the start position of the next item or location specified by the What argument. If you apply this method to the Selection object, the method moves the selection to the specified item (except for the wdGoToGrammaticalError, wdGoToProofreadingError, and wdGoToSpellingError constants).
    - `What As WdGoToItem` (required): The item where the specified range or selection is to be moved.
- `GoToPrevious(What As WdGoToItem) As Range`  
  Returns a Range object that refers to the start position of the previous item or location specified by the What argument. If applied to a Selection object, GoToPrevious moves the selection to the specified item. Range object.
    - `What As WdGoToItem` (required): The item where the specified range or selection is to be moved.
- `PasteSpecial([IconIndex As Variant], [Link As Variant], [Placement As Variant], [DisplayAsIcon As Variant], [DataType As Variant], [IconFileName As Variant], [IconLabel As Variant])`  
  Inserts the contents of the Clipboard.
    - `IconIndex As Variant` (optional): If DisplayAsIcon is True, this argument is a number that corresponds to the icon you want to use in the program file specified by IconFilename. If this argument is omitted, this method uses the first (default) icon.
    - `Link As Variant` (optional): True to create a link to the source file of the Clipboard contents. The default value is False.
    - `Placement As Variant` (optional): Can be either of the WdOLEPlacement constants.
    - `DisplayAsIcon As Variant` (optional): True to display the link as an icon. The default value is False.
    - `DataType As Variant` (optional): A format for the Clipboard contents when they are inserted into the document. WdPasteDataType.
    - `IconFileName As Variant` (optional): If DisplayAsIcon is True, this argument is the path and file name for the file in which the icon to be displayed is stored.
    - `IconLabel As Variant` (optional): If DisplayAsIcon is True, this argument is the text that appears below the icon.
- `PreviousField() As Field`  
  Selects and returns the previous field.
- `NextField() As Field`  
  Selects the next field.
- `InsertParagraphBefore()`  
  Inserts a new paragraph before the specified selection or range.
- `InsertCells([ShiftCells As Variant])`  
  Adds cells to an existing table.
    - `ShiftCells As Variant` (optional): Specifies how to insert the cells into the existing columns and rows of the table.
- `Extend([Character As Variant])`  
  Turns on extend mode, or if extend mode is already on, extends the selection to the next larger unit of text.
    - `Character As Variant` (optional): The character through which the selection is extended. This argument is case-sensitive and must evaluate to a String or an error occurs. Also, if the value of this argument is longer than a single character, Microsoft Word ignores the command entirely.
- `Shrink()`  
  Shrinks the selection to the next smaller unit of text.
- `MoveLeft([Unit As Variant], [Count As Variant], [Extend As Variant]) As Long`  
  Moves the selection to the left and returns the number of units it has been moved.
    - `Unit As Variant` (optional): The unit by which the selection is to be moved.The default value is wdCharacter.
    - `Count As Variant` (optional): The number of units the selection is to be moved. The default value is 1.
    - `Extend As Variant` (optional): Can be either wdMove or wdExtend. If wdMove is used, the selection is collapsed to the endpoint and moved to the left. If wdExtend is used, the selection is extended to the left. The default value is wdMove.
- `MoveRight([Unit As Variant], [Count As Variant], [Extend As Variant]) As Long`  
  Moves the selection to the right and returns the number of units it has been moved.
    - `Unit As Variant` (optional): The unit by which the selection is to be moved.The default value is wdCharacter.
    - `Count As Variant` (optional): The number of units the selection is to be moved. The default value is 1.
    - `Extend As Variant` (optional): Can be either wdMove or wdExtend. If wdMove is used, the selection is collapsed to the endpoint and moved to the right. If wdExtend is used, the selection is extended to the right. The default value is wdMove.
- `MoveUp([Unit As Variant], [Count As Variant], [Extend As Variant]) As Long`  
  Moves the selection up and returns the number of units that it has been moved.
    - `Unit As Variant` (optional): The unit by which to move the selection. Can be one of the following WdUnits constants: wdLine, wdParagraph, wdWindow or wdScreen. The default value is wdLine. Use the wdWindow constant for the Unit argument to move to the top or bottom of the active window. Regardless of the value of Count (greater than 1 or less than -1), the wdWindow constant moves only one unit. Use the wdScreen constant to move more than one screen.
    - `Count As Variant` (optional): The number of units the selection is to be moved. The default value is 1.
    - `Extend As Variant` (optional): Specifies whether the selection is moved or extended. Can be either wdMove or wdExtend. If wdMove is used, the selection is collapsed to the endpoint and moved up. If wdExtend is used, the selection is extended up. The default value is wdMove.
- `MoveDown([Unit As Variant], [Count As Variant], [Extend As Variant]) As Long`  
  Moves the selection down and returns the number of units it has been moved.
    - `Unit As Variant` (optional): The unit by which the selection is to be moved.The default value is wdLine.
    - `Count As Variant` (optional): The number of units the selection is to be moved. The default value is 1.
    - `Extend As Variant` (optional): Can be either wdMove or wdExtend. If wdMove is used, the selection is collapsed to the endpoint and moved down. If wdExtend is used, the selection is extended down. The default value is wdMove.
- `HomeKey([Unit As Variant], [Extend As Variant]) As Long`  
  Moves or extends the selection to the beginning of the specified unit. This method returns an integer that indicates the number of characters the selection was actually moved, or it returns 0 (zero) if the move was unsuccessful.This method corresponds to functionality of the HOME key.
    - `Unit As Variant` (optional): The unit by which the selection is to be moved or extended. The default value is wdLine.
    - `Extend As Variant` (optional): Specifies the way the selection is moved. Can be one of the WdMovementType constants. If the value of this argument is wdMove, the selection is collapsed to an insertion point and moved to the beginning of the specified unit. If it is wdExtend, the beginning of the selection is extended to the beginning of the specified unit. The default value is wdMove.
- `EndKey([Unit As Variant], [Extend As Variant]) As Long`  
  Moves or extends the selection to the end of the specified unit.
    - `Unit As Variant` (optional): The unit by which the selection is to be moved or extended. Can be a WdUnits constant. The default value is wdLine.
    - `Extend As Variant` (optional): Specifies the way the selection is moved. Can be any WdMovementType constant. If the value of this argument is wdMove, the selection is collapsed to an insertion point and moved to the end of the specified unit. If it is wdExtend, the end of the selection is extended to the end of the specified unit. The default value is wdMove.
- `EscapeKey()`  
  Cancels a mode such as extend or column select (equivalent to pressing the ESC key).
- `TypeText(Text As String)`  
  Inserts the specified text.
    - `Text As String` (required): The text to be inserted.
- `CopyFormat()`  
  Copies the character formatting of the first character in the selected text.
- `PasteFormat()`  
  Applies formatting copied with the CopyFormat method to the selection.
- `TypeParagraph()`  
  Inserts a new, blank paragraph.
- `TypeBackspace()`  
  Deletes the character preceding a collapsed selection (an insertion point).
- `NextSubdocument()`  
  Moves the selection to the next subdocument.
- `PreviousSubdocument()`  
  Moves the selection to the previous subdocument.
- `SelectColumn()`  
  Selects the column that contains the insertion point, or selects all columns that contain the selection.
- `SelectCurrentFont()`  
  Extends the selection forward until text in a different font or font size is encountered.
- `SelectCurrentAlignment()`  
  Extends the selection forward until text with a different paragraph alignment is encountered.
- `SelectCurrentSpacing()`  
  Extends the selection forward until a paragraph with different line spacing is encountered.
- `SelectCurrentIndent()`  
  Extends the selection forward until text with different left or right paragraph indents is encountered.
- `SelectCurrentTabs()`  
  Extends the selection forward until a paragraph with different tab stops is encountered.
- `SelectCurrentColor()`  
  Extends the selection forward until text with a different color is encountered.
- `CreateTextbox()`  
  Adds a default-size text box around the selection.
- `WholeStory()`  
  Expands a selection to include the entire story.
- `SelectRow()`  
  Selects the row that contains the insertion point, or selects all rows that contain the selection.
- `SplitTable()`  
  Inserts an empty paragraph above the first row in the selection.
- `InsertRows([NumRows As Variant])`  
  Inserts the specified number of new rows above the row that contains the selection. If the selection isn't in a table, an error occurs.
    - `NumRows As Variant` (optional): The number of rows to be added.
- `InsertColumns()`  
  Inserts columns to the left of the column that contains the selection.
- `InsertFormula([Formula As Variant], [NumberFormat As Variant])`  
  Inserts an = (Formula) field that contains a formula at the selection.
    - `Formula As Variant` (optional): The mathematical formula you want the = (Formula) field to evaluate. Spreadsheet-type references to table cells are valid. For example, "=SUM(A4:C4)" specifies the first three values in the fourth row. For more information about the = (Formula) field, see Field codes:= (Formula) field.
    - `NumberFormat As Variant` (optional): A format for the result of the = (Formula) field. For information about the types of formats you can apply, see Numeric Picture (\#) field switch.
- `NextRevision([Wrap As Variant]) As Revision`  
  Locates and returns the next tracked change as a Revision object.
    - `Wrap As Variant` (optional): True to continue searching for a revision at the beginning of the document when the end of the document is reached. The default value is False.
- `PreviousRevision([Wrap As Variant]) As Revision`  
  Locates and returns the previous tracked change as a Revision object.
    - `Wrap As Variant` (optional): True to continue searching for a revision at the end of the document when the beginning of the document is reached. The default value is False.
- `PasteAsNestedTable()`  
  Pastes a cell or group of cells as a nested table into the selection.
- `CreateAutoTextEntry(Name As String, StyleName As String) As AutoTextEntry`  
  Adds a new AutoTextEntry object to the AutoTextEntries collection, based on the current selection.
    - `Name As String` (required): The text the user must type to call the new AutoText entry.
    - `StyleName As String` (required): The category in which the new AutoText entry will be listed on the AutoText menu.
- `DetectLanguage()`  
  Analyzes the specified text to determine the language that it is written in.
- `SelectCell()`  
  Selects the entire cell containing the current selection.
- `InsertRowsBelow([NumRows As Variant])`  
  Inserts rows below the current selection.
- `InsertColumnsRight()`  
  Inserts columns to the right of the current selection.
- `InsertRowsAbove([NumRows As Variant])`  
  Inserts rows above the current selection.
- `RtlRun()`  
  Sets the reading order and alignment of the specified run to right-to-left.
- `LtrRun()`  
  Sets the reading order and alignment of the specified run to left-to-right.
- `BoldRun()`  
  Adds the bold character format to or removes it from the current run.
- `ItalicRun()`  
  Adds the italic character format to or removes it from the current run.
- `RtlPara()`  
  Sets the reading order and alignment of the specified paragraphs to right-to-left.
- `LtrPara()`  
  Sets the reading order and alignment of the specified paragraphs to left-to-right.
- `InsertDateTime([DateTimeFormat As Variant], [InsertAsField As Variant], [InsertAsFullWidth As Variant], [DateLanguage As Variant], [CalendarType As Variant])`  
  Inserts the current date or time, or both, either as text or as a TIME field.
    - `DateTimeFormat As Variant` (optional): The format to be used for displaying the date or time, or both. If this argument is omitted, Microsoft Word uses the short-date style from the Windows Control Panel (Regional Settings icon). See Date/Time data type for format symbols.
    - `InsertAsField As Variant` (optional): True to insert the specified information as a TIME field. The default value is True.
    - `InsertAsFullWidth As Variant` (optional): True to insert the specified information as double-byte digits. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `DateLanguage As Variant` (optional): Sets the language in which to display the date or time. Can be either of the WdDateLanguage constants. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `CalendarType As Variant` (optional): Sets the calendar type to use when displaying the date or time. Can be either of the WdCalendarTypeBi constants. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
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
- `ClearFormatting()`  
  Removes text and paragraph formatting from a selection.
- `PasteAppendTable()`  
  Merges pasted cells into an existing table by inserting the pasted rows between the selected rows. No cells are overwritten.
- `ToggleCharacterCode()`  
  Switches a selection between a Unicode character and its corresponding hexadecimal value.
- `PasteAndFormat(Type As WdRecoveryType)`  
  Pastes the selected table cells and formats them as specified.
    - `Type As WdRecoveryType` (required): The type of formatting to use when pasting the selected table cells.
- `PasteExcelTable(LinkedToExcel As Boolean, WordFormatting As Boolean, RTF As Boolean)`  
  Pastes and formats a Microsoft Excel table.
    - `LinkedToExcel As Boolean` (required): True links the pasted table to the original Excel file so that changes made to the Excel file are reflected in Microsoft Word.
    - `WordFormatting As Boolean` (required): True formats the table using the formatting in the Word document. False formats the table according to the original Excel file.
    - `RTF As Boolean` (required): True pastes the Excel table using Rich Text Format (RTF). False pastes the Excel table as HTML.
- `ShrinkDiscontiguousSelection()`  
  Cancels the selection of all but the most recently selected text when a selection contains multiple, unconnected selections.
- `InsertStyleSeparator()`  
  Inserts a special hidden paragraph mark that allows Microsoft Word to join paragraphs formatted using different paragraph styles, so lead-in headings can be inserted into a table of contents.
- `Sort([ExcludeHeader As Variant], [FieldNumber As Variant], [SortFieldType As Variant], [SortOrder As Variant], [FieldNumber2 As Variant], [SortFieldType2 As Variant], [SortOrder2 As Variant], [FieldNumber3 As Variant], [SortFieldType3 As Variant], [SortOrder3 As Variant], [SortColumn As Variant], [Separator As Variant], [CaseSensitive As Variant], [BidiSort As Variant], [IgnoreThe As Variant], [IgnoreKashida As Variant], [IgnoreDiacritics As Variant], [IgnoreHe As Variant], [LanguageID As Variant], [SubFieldNumber As Variant], [SubFieldNumber2 As Variant], [SubFieldNumber3 As Variant])`  
  Sorts the paragraphs in the specified selection.
    - `ExcludeHeader As Variant` (optional): True to exclude the first row or paragraph header from the sort operation. The default value is False.
    - `FieldNumber As Variant` (optional): The first field by which to sort.
    - `SortFieldType As Variant` (optional): The sort type for FieldNumber. Can be one of the WdSortFieldType constants. The default value is wdSortFieldAlphanumeric. Some of the WdSortFieldType constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `SortOrder As Variant` (optional): The sorting order to use when sorting FieldNumber2. Can be one WdSortOrder constant.The default value is wdSortOrderAscending.
    - `FieldNumber2 As Variant` (optional): The second field by which to sort.
    - `SortFieldType2 As Variant` (optional): The sort type for FieldNumber2. Can be one of the WdSortFieldType constants. The default value is wdSortFieldAlphanumeric. Some of the WdSortFieldType constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `FieldNumber3 As Variant` (optional): The third field by which to sort.
    - `SortFieldType3 As Variant` (optional): The sort type for FieldNumber3. Can be one of the WdSortFieldType constants. The default value is wdSortFieldAlphanumeric. Some of the WdSortFieldType constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `SortOrder3 As Variant` (optional): The sorting order to use when sorting FieldNumber3. Can be one WdSortOrder constant.The default value is wdSortOrderAscending.
    - `Separator As Variant` (optional): The type of field separator.
    - `CaseSensitive As Variant` (optional): True to sort with case sensitivity. The default value is False.
    - `BidiSort As Variant` (optional): True to sort based on right-to-left language rules. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `IgnoreThe As Variant` (optional): True to ignore the Arabic character alef lam when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `IgnoreKashida As Variant` (optional): True to ignore kashidas when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `IgnoreDiacritics As Variant` (optional): True to ignore bidirectional control characters when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `IgnoreHe As Variant` (optional): True to ignore the Hebrew character he when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `LanguageID As Variant` (optional): Specifies the sorting language. Can be one of the WdLanguageID constants.
    - `SubFieldNumber As Variant` (optional): A secondary field number by which to sort.
    - `SubFieldNumber2 As Variant` (optional): A secondary field number by which to sort.
    - `SubFieldNumber3 As Variant` (optional): A secondary field number by which to sort.
- `GoToEditableRange([EditorID As Variant]) As Range`  
  Returns a Range object that represents an area of a document that can be modified by the specified user or group of users.
    - `EditorID As Variant` (optional): Can be either a String that represents the user's email alias (if in the same domain), an email address, or a WdEditorType constant that represents a group of users. If omitted, selects all ranges for which all users have permissions to edit.
- `InsertXML(XML As String, [Transform As Variant])`  
  Inserts the specified XML into the document at the cursor, replacing any selected text.
    - `XML As String` (required): Specifies the XML to insert. This can be any valid custom XML.
    - `Transform As Variant` (optional): Specifies the XML Transformation (XSLT) used to transform the XML. If omitted, the XML is inserted as custom XML without applying a transform.
- `InsertCaption(Label As Variant, [Title As Variant], [TitleAutoText As Variant], [Position As Variant], [ExcludeLabel As Variant])`  
  Inserts a caption immediately preceding or following the specified selection.
    - `Label As Variant` (required): The caption label to be inserted. Can be a String or one of the WdCaptionLabelID constants. If the label has not yet been defined, an error occurs. Use the Add method with the CaptionLabels object to define new caption labels.
    - `Title As Variant` (optional): The string to be inserted immediately following the label in the caption (ignored if TitleAutoText is specified).
    - `TitleAutoText As Variant` (optional): The AutoText entry whose contents you want to insert immediately following the label in the caption (overrides any text specified by Title).
    - `Position As Variant` (optional): Specifies whether the caption will be inserted above or below the selection. Can be one of the WdCaptionPosition constants.
    - `ExcludeLabel As Variant` (optional): True does not include the text label, as defined in the Label parameter. False includes the specified label.
- `InsertCrossReference(ReferenceType As Variant, ReferenceKind As WdReferenceKind, ReferenceItem As Variant, [InsertAsHyperlink As Variant], [IncludePosition As Variant], [SeparateNumbers As Variant], [SeparatorString As Variant])`  
  Inserts a cross-reference to a heading, bookmark, footnote, or endnote, or to an item for which a caption label is defined (for example, an equation, figure, or table).
    - `ReferenceType As Variant` (required): The type of item for which a cross-reference is to be inserted. Can be any WdReferenceType or WdCaptionLabelID constant or a user defined caption label.
    - `ReferenceKind As WdReferenceKind` (required): The information to be included in the cross-reference.
    - `ReferenceItem As Variant` (required): If ReferenceType is wdRefTypeBookmark, this argument specifies a bookmark name. For all other ReferenceType values, this argument specifies the item number or name in the Reference type box in the Cross-reference dialog box. Use the GetCrossReferenceItems method to return a list of item names that can be used with this argument.
    - `InsertAsHyperlink As Variant` (optional): True to insert the cross-reference as a hyperlink.
    - `IncludePosition As Variant` (optional): True to insert "above" or "below," depending on the location of the reference item in relation to the cross-reference.
    - `SeparateNumbers As Variant` (optional): True to use a separator to separate the numbers from the associated text. (Use only if the ReferenceType parameter is set to wdRefTypeNumberedItem and the ReferenceKind parameter is set to wdNumberFullContext.)
    - `SeparatorString As Variant` (optional): Specifies the string to use as a separator if the SeparateNumbers parameter is set to True.
- `ClearParagraphStyle()`  
  Removes paragraph formatting that has been applied through paragraph styles from the selected text.
- `ClearCharacterAllFormatting()`  
  Removes all character formatting (formatting applied either through character styles or manually applied formatting) from the selected text.
- `ClearCharacterStyle()`  
  Removes character formatting that has been applied through character styles from the selected text.
- `ClearCharacterDirectFormatting()`  
  Removes character formatting (formatting that has been applied manually using the buttons on the ribbon or through the dialog boxes) from the selected text.
- `ExportAsFixedFormat(OutputFileName As String, ExportFormat As WdExportFormat, [OpenAfterExport As Boolean], [OptimizeFor As WdExportOptimizeFor], [ExportCurrentPage As Boolean], [Item As WdExportItem], [IncludeDocProps As Boolean], [KeepIRM As Boolean], [CreateBookmarks As WdExportCreateBookmarks], [DocStructureTags As Boolean], [BitmapMissingFonts As Boolean], [UseISO19005_1 As Boolean], [FixedFormatExtClassPtr As Variant])`  
  Saves the current selection as PDF or XPS format.
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
    - `FixedFormatExtClassPtr As Variant` (optional): Specifies a pointer to an add-in that allows calls to an alternate implementation of code. The alternate implementation of code interprets the EMF and EMF+ page descriptions that are generated by the applications to make their own PDF or XPS. For more information, see Extend the fixed-format export feature in Word Automation Services.
- `ReadingModeGrowFont()`  
  Increases the size of the displayed text one point size when the document is displayed in Reading mode.
- `ReadingModeShrinkFont()`  
  Decreases the size of the displayed text one point size when the document is displayed in Reading mode.
- `ClearParagraphAllFormatting()`  
  Removes all paragraph formatting (formatting applied either through paragraph styles or manually applied formatting) from the selected text.
- `ClearParagraphDirectFormatting()`  
  Removes paragraph formatting that has been applied manually (using the buttons on the ribbon or through the dialog boxes) from the selected text.
- `InsertNewPage()`  
  Inserts a new page at the position of the Insertion Point.
- `SortByHeadings([SortFieldType As Variant], [SortOrder As Variant], [CaseSensitive As Variant], [BidiSort As Variant], [IgnoreThe As Variant], [IgnoreKashida As Variant], [IgnoreDiacritics As Variant], [IgnoreHe As Variant], [LanguageID As Variant])`  
  Sorts the headings in the specified selection.
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
  Saves the current selection as PDF or XPS format.
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
  Saves the current selection as PDF or XPS format.
    - `OutputFileName As String` (required): The path and file name of the new PDF or XPS file.
    - `ExportFormat As WdExportFormat` (required): Specifies either PDF or XPS format.
    - `OpenAfterExport As Boolean` (optional): Opens the new file after exporting the contents.
    - `OptimizeFor As WdExportOptimizeFor` (optional): Specifies whether to optimize for screen or print.
    - `ExportCurrentPage As Boolean` (optional): Specifies whether to export the current page. True exports the entire page. False exports only the current selection.
    - `Item As WdExportItem` (optional): Specifies whether the export process includes text only or includes text with markup.
    - `IncludeDocProps As Boolean` (optional): Specifies whether to include document properties in the newly exported file.
    - `KeepIRM As Boolean` (optional): Specifies whether to copy IRM permissions to an XPS document if the source document has IRM protections.</br></br>If _ExportFormat_ is _wdExportFormatPDF_, this flag also specifies whether to copy labels to the PDF.</br></br>Default value is True.
    - `CreateBookmarks As WdExportCreateBookmarks` (optional): Specifies whether to export bookmarks and the type of bookmarks to export.
    - `DocStructureTags As Boolean` (optional): Specifies whether to include extra data to help screen readers, for example information about the flow and logical organization of the content. Default value is True.
    - `BitmapMissingFonts As Boolean` (optional): Specifies whether to include a bitmap of the text. Set this parameter to True when font licenses don't permit a font to be embedded in the PDF file. If False, the font is referenced, and the viewer's computer substitutes an appropriate font if the authored one is not available. Default value is True.
    - `OptimizeForImageQuality As Boolean` (optional): Specifies whether to downsample images or keep their original quality. If True, the resulting files will have better image quality but may be larger. Default value is False.
    - `ImproveExportTagging As Boolean` (optional): Specifies whether to enable improved accessbility tagging. For more information, see the Remarks section. Default value is False.
    - `FixedFormatExtClassPtr As Variant` (optional): Specifies a pointer to an add-in that allows calls to an alternate implementation of code. The alternate implementation of code interprets the EMF and EMF+ page descriptions that are generated by the applications to make their own PDF or XPS.
