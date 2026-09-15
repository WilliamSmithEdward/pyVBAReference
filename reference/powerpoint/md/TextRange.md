# TextRange

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {9149348F-5A91-11CF-8700-00AA0060263B}  

Contains the text that's attached to a shape, and properties and methods for manipulating the text.

**Remarks:** The following examples describe how to: - Return the text range in any shape you specify. - Return a text range from the selection. - Return particular characters, words, lines, sentences, or paragraphs from a text range. - Find and replace text in a text range. - Insert text, the date and time, or the slide number into a text range. - Position the cursor wherever you want in a text range.

**Example:**

```vba
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140) _

    .TextFrame.TextRange.Text = "Here is some test text"
```

## Properties (15)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `ActionSettings As ActionSettings  (read-only)`  
  Returns an ActionSettings object that contains information about what action occurs when the user clicks or moves the mouse over the specified shape or text range during a slide show. Read-only.
- `Start As Long  (read-only)`  
  Returns the position of the first character in the specified text range relative to the first character in the shape that contains the text. Read-only.
- `Length As Long  (read-only)`  
  Returns the length of the specified text range, in characters. Read-only.
- `BoundLeft As Single  (read-only)`  
  Returns the distance (in points) from the left edge of the text bounding box for the specified text frame to the left edge of the slide. Read-only.
- `BoundTop As Single  (read-only)`  
  Returns the distance (in points) from the top of the of the text bounding box for the specified text frame to the top of the slide. Read-only.
- `BoundWidth As Single  (read-only)`  
  Returns the width (in points) of the text bounding box for the specified text frame. Read-only.
- `BoundHeight As Single  (read-only)`  
  Returns the height (in points) of the text bounding box for the specified text frame. Read-only.
- `Text As String  (read/write)`  
  Returns or sets a String that represents the text contained in the specified object. Read/write.
- `Font As Font  (read-only)`  
  Returns a Font object that represents character formatting. Read-only.
- `ParagraphFormat As ParagraphFormat  (read-only)`  
  Returns a ParagraphFormat object that represents paragraph formatting for the specified text. Read-only.
- `IndentLevel As Long  (read/write)`  
  Returns or sets the indent level for the specified text as an integer from 1 to 5, where 1 indicates a first-level paragraph with no indentation. Read/write.
- `LanguageID As MsoLanguageID  (read/write)`  
  Returns or sets the language for the specified text range. Read/write.

## Methods (26)

- `Paragraphs([Start As Long], [Length As Long]) As TextRange`  
  Returns a TextRange object that represents the specified subset of text paragraphs.
    - `Start As Long` (optional): The first paragraph in the returned range.
    - `Length As Long` (optional): The number of paragraphs to be returned.
- `Sentences([Start As Long], [Length As Long]) As TextRange`  
  Returns a TextRange object that represents the specified subset of text sentences.
    - `Start As Long` (optional): The first sentence in the returned range.
    - `Length As Long` (optional): The number of sentences to be returned.
- `Words([Start As Long], [Length As Long]) As TextRange`  
  Returns a TextRange object that represents the specified subset of text words.
    - `Start As Long` (optional): The first word in the returned range.
    - `Length As Long` (optional): The number of words to be returned.
- `Characters([Start As Long], [Length As Long]) As TextRange`  
  Returns a TextRange object that represents the specified subset of text characters. For information about counting or looping through the characters in a text range, see the TextRange object.
    - `Start As Long` (optional): The first character in the returned range.
    - `Length As Long` (optional): The number of characters to be returned.
- `Lines([Start As Long], [Length As Long]) As TextRange`  
  Returns a TextRange object that represents the specified subset of text lines. For information about counting or looping through the lines in a text range, see the TextRange object.
    - `Start As Long` (optional): The first line in the returned range.
    - `Length As Long` (optional): The number of lines to be returned.
- `Runs([Start As Long], [Length As Long]) As TextRange`  
  Returns a TextRange object that represents the specified subset of text runs. A text run consists of a range of characters that share the same font attributes.
    - `Start As Long` (optional): The first run in the returned range.
    - `Length As Long` (optional): The number of runs to be returned.
- `TrimText() As TextRange`  
  Returns a TextRange object that represents the specified text minus any trailing spaces.
- `InsertAfter([NewText As String]) As TextRange`  
  Appends a string to the end of the specified text range. Returns a TextRange object that represents the appended text. When used without an argument, this method returns a zero-length string at the end of the specified range.
    - `NewText As String` (optional): The text to be inserted. The default value is an empty string.
- `InsertBefore([NewText As String]) As TextRange`  
  Appends a string to the beginning of the specified text range. Returns a TextRange object that represents the appended text. When used without an argument, this method returns a zero-length string at the end of the specified range.
    - `NewText As String` (optional): The text to be appended. The default value is an empty string.
- `InsertDateTime(DateTimeFormat As PpDateTimeFormat, [InsertAsField As MsoTriState]) As TextRange`  
  Inserts the date and time in the specified text range. Returns a TextRange object that represents the inserted text.
    - `DateTimeFormat As PpDateTimeFormat` (required): A format for the date and time.
    - `InsertAsField As MsoTriState` (optional): Determines whether the inserted date and time will be updated each time the presentation is opened.
- `InsertSlideNumber() As TextRange`  
  Inserts the slide number of the current slide into the specified text range. Returns a TextRange object that represents the slide number.
- `InsertSymbol(FontName As String, CharNumber As Long, [Unicode As MsoTriState]) As TextRange`  
  Returns a TextRange object that represents a symbol inserted into the specified text range.
    - `FontName As String` (required): The font name.
    - `CharNumber As Long` (required): The Unicode or ASCII character number.
    - `Unicode As MsoTriState` (optional): Specifies whether the CharNumber argument represents an ASCII or Unicode character.
- `Select()`  
  Selects the specified object.
- `Cut()`  
  Deletes the specified object and places it on the Clipboard.
- `Copy()`  
  Copies the specified object to the Clipboard.
- `Delete()`  
  Deletes the specified TextRange object.
- `Paste() As TextRange`  
  Pastes the text on the Clipboard into the specified text range, and returns a TextRange object that represents the pasted text.
- `ChangeCase(Type As PpChangeCase)`  
  Changes the case of the specified text.
    - `Type As PpChangeCase` (required): Specifies the way the case will be changed.
- `AddPeriods()`  
  Adds a period at the end of each paragraph in the specified text.
- `RemovePeriods()`  
  Removes the period at the end of each paragraph in the specified text.
- `Find(FindWhat As String, [After As Long], [MatchCase As MsoTriState], [WholeWords As MsoTriState]) As TextRange`  
  Finds the specified text in a text range, and returns a TextRange object that represents the first text range where the text is found. Returns Nothing if no match is found.
    - `FindWhat As String` (required): The text to search for.
    - `After As Long` (optional): The position of the character (in the specified text range) after which you want to search for the next occurrence of FindWhat. For example, if you want to search from the fifth character of the text range, specify 4 for After. If this argument is omitted, the first character of the text range is used as the starting point for the search.
    - `MatchCase As MsoTriState` (optional): msoTrue for the search to distinguish between uppercase and lowercase characters.
    - `WholeWords As MsoTriState` (optional): msoTrue for the search to find only whole words and not parts of larger words as well.
- `Replace(FindWhat As String, ReplaceWhat As String, [After As Long], [MatchCase As MsoTriState], [WholeWords As MsoTriState]) As TextRange`  
  Finds specific text in a text range, replaces the found text with a specified string, and returns a TextRange object that represents the first occurrence of the found text. Returns Nothing if no match is found.
    - `FindWhat As String` (required): The text to search for.
    - `ReplaceWhat As String` (required): The text you want to replace the found text with.
    - `After As Long` (optional): The position of the character (in the specified text range) after which you want to search for the next occurrence of FindWhat. For example, if you want to search from the fifth character of the text range, specify 4 for After. If this argument is omitted, the first character of the text range is used as the starting point for the search.
    - `MatchCase As MsoTriState` (optional): Determines whether a distinction is made on the basis of case.
    - `WholeWords As MsoTriState` (optional): Determines whether only whole words are found.
- `RotatedBounds(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, X3 As Single, Y3 As Single, x4 As Single, y4 As Single)`  
  Returns the coordinates of the vertices of the text bounding box for the specified text range.
- `RtlRun()`  
  Sets the direction of text in a text range to read from right to left.
- `LtrRun()`  
  Sets the direction of text in a text range to read from left to right.
- `PasteSpecial([DataType As PpPasteDataType], [DisplayAsIcon As MsoTriState], [IconFileName As String], [IconIndex As Long], [IconLabel As String], [Link As MsoTriState]) As TextRange`  
  Replaces the text range with the contents of the Clipboard in the format specified.
    - `DataType As PpPasteDataType` (optional): A format for the Clipboard contents when they're inserted into the document. The default value varies, depending on the contents in the Clipboard. An error occurs if the specified data type in the DataType argument is not supported by the clipboard contents.
    - `DisplayAsIcon As MsoTriState` (optional): MsoTrue to display the embedded object (or link) as an icon.
    - `IconFileName As String` (optional): If DisplayAsIcon is set to msoTrue, this argument is the path and file name for the file in which the icon to be displayed is stored. If DisplayAsIcon is set to msoFalse, this argument is ignored.
    - `IconIndex As Long` (optional): If DisplayAsIcon is set to msoTrue, this argument is a number that corresponds to the icon you want to use in the program file specified by IconFilename. For example, 0 (zero) corresponds to the first icon, 1 corresponds to the second icon. If this argument is omitted, the first (default) icon is used. If DisplayAsIcon is set to msoFalse, then this argument is ignored. If IconIndex is outside the valid range, then the default icon (index 0) is used.
    - `IconLabel As String` (optional): If DisplayAsIcon is set to msoTrue, this argument is the text that appears below the icon. If this label is missing, Microsoft PowerPoint generates an icon label based on the Clipboard contents. If DisplayAsIcon is set to msoFalse, then this argument is ignored.
    - `Link As MsoTriState` (optional): Determines whether to create a link to the source file of the Clipboard contents. An error occurs if the Clipboard contents don't support a link.
