# Footnotes

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020942-0000-0000-C000-000000000046}  

A collection of Footnote objects that represent all the footnotes in a selection, range, or document.

**Remarks:** Use the Footnotes property to return the Footnotes collection. The following example changes all of the footnotes in the active document to endnotes. Use the Add method to add a footnote to the Footnotes collection. The following example adds a footnote immediately after the selection. Use Footnotes (index), where index is the index number, to return a single Footnote object. The index number represents the position of the footnote in the selection, range, or document. The following example applies red formatting to the first footnote in the selection.

## Properties (12)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of footnotes in the collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Footnotes object.
- `Location As WdFootnoteLocation  (read/write)`  
  Returns or sets the position of all footnotes. Read/write WdFootnoteLocation.
- `NumberStyle As WdNoteNumberStyle  (read/write)`  
  Returns or sets the number style for the footnotes. Read/write WdNoteNumberStyle.
- `StartingNumber As Long  (read/write)`  
  Returns or sets the starting note number, line number, or page number. Read/write Long.
- `NumberingRule As WdNumberingRule  (read/write)`  
  Returns or sets the way footnotes or endnotes are numbered after page breaks or section breaks. Read/write WdNumberingRule.
- `Separator As Range  (read-only)`  
  Returns a Range object that represents the footnote separator.
- `ContinuationSeparator As Range  (read-only)`  
  Returns a Range object that represents the footnote continuation separator. Read-only.
- `ContinuationNotice As Range  (read-only)`  
  Returns a Range object that represents the footnote continuation notice. Read-only.

## Methods (7)

- `Item(Index As Long) As Footnote`  
  Returns an individual Footnote object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
- `Add(Range As Range, [Reference As Variant], [Text As Variant]) As Footnote`  
  Returns a Footnote object that represents a footnote added to a range.
    - `Range As Range` (required): The range marked for the endnote or footnote. This can be a collapsed range.
    - `Reference As Variant` (optional): The text for the custom reference mark. If this argument is omitted, Microsoft Word inserts an automatically-numbered reference mark.
    - `Text As Variant` (optional): The text of the endnote or footnote.
- `Convert()`  
  Converts endnotes to footnotes, or vice versa.
- `SwapWithEndnotes()`  
  Converts all footnotes in a document to endnotes and vice versa.To convert a range of footnotes to endnotes, use the Convert method.
- `ResetSeparator()`  
  Resets the footnote separator to the default separator.
- `ResetContinuationSeparator()`  
  Resets the footnote or endnote continuation separator to the default separator.
- `ResetContinuationNotice()`  
  Resets the footnote or endnote continuation notice to the default notice.
