# TextRange2

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0397-0000-0000-C000-000000000046}  

Represents the text frame in a Shape or ShapeRange object.

**Remarks:** This object contains the text in the text frame as well as the properties and methods that control the alignment and anchoring of the text frame. Use the TextFrame2 property to return a TextFrame2 object.

**Example:**

```vba
Set myDocument = Worksheets(1)
With myDocument.Shapes.AddShape(msoShapeRectangle, _
 0, 0, 250, 140).TextFrame2
 .TextRange.Text = "Here is some test text"
 .MarginBottom = 10
 .MarginLeft = 10
 .MarginRight = 10
 .MarginTop = 10
End With
```

## Properties (22)

- `Application As Object  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the current instance of the Microsoft Office application. When used with an object qualifier, this property returns an Application object that represents the creator of the TextRange2 object. When used with an OLE Automation object, it returns the object's application. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the TextRange2 object was created. Read-only.
- `Text As String  (read/write)`  
  Gets or sets a String value that represents the text in a text range. Read/write.
- `Count As Long  (read-only)`  
  Gets a Long indicating the number of items in the TextRange2 collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`
- `Parent As Object  (read-only)`  
  Gets the Parent object for the TextRange2 object. Read-only.
- `Paragraphs As TextRange2  (read-only)`  
  Gets a TextRange2 object that represents the specified subset of text paragraphs. Read-only.
- `Sentences As TextRange2  (read-only)`  
  Returns a TextRange2 object that represents the specified subset of text sentences. Read-only.
- `Words As TextRange2  (read-only)`  
  Gets a TextRange2 object that represents the specified subset of text words. Read-only.
- `Characters As TextRange2  (read-only)`  
  Read-only.
- `Lines As TextRange2  (read-only)`  
  Returns a TextRange2 object that represents the specified subset of text lines. Read-only.
- `Runs As TextRange2  (read-only)`  
  Gets a TextRange2 object that represents the specified subset of text runs. A text run consists of a range of characters that share the same font attributes. Read-only.
- `ParagraphFormat As ParagraphFormat2  (read-only)`  
  Returns a ParagraphFormat object that represents paragraph formatting for the specified text. Read-only.
- `Font As Font2  (read-only)`  
  Returns a Font object that represents character formatting for the TextRange2 object. Read-only.
- `Length As Long  (read-only)`  
  Gets a Long that represents the length of a text range. Read-only.
- `Start As Long  (read-only)`  
  Gets a Long value indicating the starting point of the specified text range. Read-only.
- `BoundLeft As Single  (read-only)`  
  Gets the left coordinate, in points, of the text bounding box for the specified text. Read-only.
- `BoundTop As Single  (read-only)`  
  Gets the top coordinate, in points, of the text bounding box for the specified text. Read-only.
- `BoundWidth As Single  (read-only)`  
  Gets the width, in points, of the text bounding box for the specified text. Read-only.
- `BoundHeight As Single  (read-only)`  
  Gets the height, in points, of the text bounding box for the specified text. Read-only.
- `LanguageID As MsoLanguageID  (read/write)`  
  Gets or sets the MsoLanguageID value of the TextRange2 object. Read/write.
- `MathZones As TextRange2  (read-only)`  
  Sets the starting point and length of a math zone within a text range. Read-only.

## Methods (20)

- `Item(Index As Variant) As TextRange2`  
  Gets the range of text specified by the index number from the TextRange2 object.
    - `Index As Variant` (required): The index number of the text range.
- `TrimText() As TextRange2`  
  Returns a TextRange2 object that represents the specified text that has the whitespace removed.
- `InsertAfter([NewText As String]) As TextRange2`  
  Inserts text to the right of the existing text in the TextRange2 object.
    - `NewText As String` (optional): Contains the text to be inserted.
- `InsertBefore([NewText As String]) As TextRange2`  
  Inserts text to the left of the existing text in the TextRange2 object.
    - `NewText As String` (optional): Contains the text to be inserted.
- `InsertSymbol(FontName As String, CharNumber As Long, [Unicode As MsoTriState]) As TextRange2`  
  Inserts a symbol from the specified font set into the range of text represented by the TextRange2 object.
    - `FontName As String` (required): The name of the font set.
    - `CharNumber As Long` (required): The number of the symbol.
    - `Unicode As MsoTriState` (optional): Indicates whether the value of the symbol is specified as a unicode value.
- `Select()`  
  Selects the TextRange2 object.
- `Cut()`  
  Removes a portion or all of the text from a range of text.
- `Copy()`  
  Copies a TextRange2 object.
- `Delete()`  
  Deletes a TextRange2 object.
- `Paste() As TextRange2`  
  Pastes the contents of the Clipboard into the TextRange2 object.
- `PasteSpecial(Format As MsoClipboardFormat) As TextRange2`  
  Replaces the text range with the contents of the Clipboard in the format specified. If the paste succeeds, this method returns a TextRange2 object, including the text range that was pasted.
    - `Format As MsoClipboardFormat` (required): Determines the format for the Clipboard contents when they're inserted into the document.
- `ChangeCase(Type As MsoTextChangeCase)`  
  Changes the case of a TextRange2 object to one of the values in the MsoTextChangeCase enumeration.
    - `Type As MsoTextChangeCase` (required): Specifies the type of change to make to the text.
- `AddPeriods()`  
  Adds period (.) punctuation to the right side of the text contained in a TextRange2 object for left-to-right languages and on the left side for right-to-left languages.
- `RemovePeriods()`  
  Removes all period (.) punctuation from the text in the TextRange2 object.
- `Find(FindWhat As String, [After As Long], [MatchCase As MsoTriState], [WholeWords As MsoTriState]) As TextRange2`  
  Searches a TextRange2 object for a subset of text.
    - `FindWhat As String` (required): Contains the text to find.
    - `After As Long` (optional): Specifies the point in the text range to start the search.
    - `MatchCase As MsoTriState` (optional): Specifies if the target text must exactly match the case of the search text.
    - `WholeWords As MsoTriState` (optional): Specifies that only whole words will be searched.
- `Replace(FindWhat As String, ReplaceWhat As String, [After As Long], [MatchCase As MsoTriState], [WholeWords As MsoTriState]) As TextRange2`  
  Finds specific text in a text range, replaces the found text with a specified string, and returns a TextRange2 object that represents the first occurrence of the found text. Returns Nothing if no match is found.
    - `FindWhat As String` (required): The text to search for.
    - `ReplaceWhat As String` (required): The text you want to replace the found text with.
    - `After As Long` (optional): The position of the character (in the specified text range) after which you want to search for the next occurrence of FindWhat. For example, if you want to search from the fifth character of the text range, specify 4 for After. If this argument is omitted, the first character of the text range is used as the starting point for the search.
    - `MatchCase As MsoTriState` (optional): Determines whether a distinction is made on the basis of case.
    - `WholeWords As MsoTriState` (optional): Determines whether only whole words are searched.
- `RotatedBounds(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, X3 As Single, Y3 As Single, x4 As Single, y4 As Single)`  
  Gets the coordinates of the vertices of the text bounding box for the specified text range. Read-only.
    - `X1 As Single` (required): Returns the position (in points) of the _X_ coordinate of the first vertex of the bounding box for the text within the specified text range.
    - `Y1 As Single` (required): Returns the position (in points) of the _Y_ coordinate of the first vertex of the bounding box for the text within the specified text range.
    - `X2 As Single` (required): Returns the position (in points) of the _X_ coordinate of the second vertex of the bounding box for the text within the specified text range.
    - `Y2 As Single` (required): Returns the position (in points) of the _Y_ coordinate of the second vertex of the bounding box for the text within the specified text range.
    - `X3 As Single` (required): Returns the position (in points) of the _X_ coordinate of the third vertex of the bounding box for the text within the specified text range.
    - `Y3 As Single` (required): Returns the position (in points) of the _Y_ coordinate of the third vertex of the bounding box for the text within the specified text range.
    - `x4 As Single` (required): Returns the position (in points) of the _X_ coordinate of the fourth vertex of the bounding box for the text within the specified text range.
    - `y4 As Single` (required): Returns the position (in points) of the _Y_ coordinate of the fourth vertex of the bounding box for the text within the specified text range.
- `RtlRun()`  
  Returns a TextRange2 object that represents the specified subset of right-to-left text runs. A text run consists of a range of characters that share the same font attributes.
- `LtrRun()`  
  Returns a TextRange2 object that represents the specified subset of left-to-right text runs. A text run consists of a range of characters that share the same font attributes.
- `InsertChartField(ChartFieldType As MsoChartFieldType, [Formula As String], [Position As Long]) As TextRange2`  
  Inserts a field into the body of a data label in a chart.
    - `ChartFieldType As MsoChartFieldType` (required): Specifies the type of chart field to insert into a data label.
    - `Formula As String` (optional): Specifies a cell (or range) if the msoChartFieldFormula constant is passed in for the _ChartFieldType_ parameter.
    - `Position As Long` (optional): Specifies the character position where the chart field is inserted. The default is to append the field to the end of the text. If the position value is out of range, the default is used.
