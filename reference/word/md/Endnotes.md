# Endnotes

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020941-0000-0000-C000-000000000046}  

A collection of Endnote objects that represents all the endnotes in a selection, range, or document.

**Remarks:** Use the Endnotes property to return the Endnotes collection. The following example sets the location of endnotes in the active document. Use the Add method to add an endnote to the Endnotes collection. The following example adds an endnote immediately after the selection. Use Endnotes (Index), where Index is the index number, to return a single Endnote object. The index number represents the position of the endnote in a selection, range, or document. The following example applies red formatting to the first endnote in the selection.

## Properties (12)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of endnotes in the collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Endnotes object.
- `Location As WdEndnoteLocation  (read/write)`  
  Returns or sets the position of all endnotes. Read/write WdEndnoteLocation. .
- `NumberStyle As WdNoteNumberStyle  (read/write)`  
  Returns or sets the number style. Read/write WdNoteNumberStyle.
- `StartingNumber As Long  (read/write)`  
  Returns or sets the starting note number, line number, or page number. Read/write Long.
- `NumberingRule As WdNumberingRule  (read/write)`  
  Returns or sets the way endnotes are numbered after page breaks or section breaks. Read/write WdNumberingRule.
- `Separator As Range  (read-only)`  
  Returns a Range object that represents the endnote separator.
- `ContinuationSeparator As Range  (read-only)`  
  Returns a Range object that represents the endnote continuation separator. Read-only.
- `ContinuationNotice As Range  (read-only)`  
  Returns a Range object that represents the endnote continuation notice. Read-only.

## Methods (7)

- `Item(Index As Long) As Endnote`  
  Returns an individual Endnote object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
- `Add(Range As Range, [Reference As Variant], [Text As Variant]) As Endnote`  
  Returns an Endnote object that represents an endnote added to a range.
    - `Range As Range` (required): The range marked for the endnote or footnote. This can be a collapsed range.
    - `Reference As Variant` (optional): The text for the custom reference mark. If this argument is omitted, Microsoft Word inserts an automatically-numbered reference mark.
    - `Text As Variant` (optional): The text of the endnote or footnote.
- `Convert()`  
  Converts endnotes to footnotes.
- `SwapWithFootnotes()`  
  Converts all endnotes in a document to footnotes and vice versa.
- `ResetSeparator()`  
  Resets the endnote separator to the default separator.
- `ResetContinuationSeparator()`  
  Resets the endnote continuation separator to the default separator.
- `ResetContinuationNotice()`  
  Resets the endnote continuation notice to the default notice.
