# EndnoteOptions

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {BF043168-F4DE-4E7C-B206-741A8B3EF71A}  

Represents the properties assigned to a range or selection of endnotes in a document.

**Remarks:** Use the EndnoteOptions property of the Range or Selection object to return an EndnoteOptions object. Using the EndnoteOptions object, you can assign different endnote properties to different areas of a document. For example, you may want endnotes in the introduction of a long document to be displayed as lowercase Roman numerals, while in the rest of your document they are displayed as Arabic numerals. The following example uses the NumberingRule, NumberStyle, and StartingNumber properties to format the endnotes in the first section ofthe active document.

## Properties (7)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified EndnoteOptions object.
- `Location As WdEndnoteLocation  (read/write)`  
  Returns or sets the position of all endnotes. Read/write WdEndnoteLocation.
- `NumberStyle As WdNoteNumberStyle  (read/write)`  
  Returns or sets the number style for the endnotes. Read/write WdNoteNumberStyle.
- `StartingNumber As Long  (read/write)`  
  Returns or sets the starting endnote number. Read/write Long.
- `NumberingRule As WdNumberingRule  (read/write)`  
  Returns or sets the way footnotes or endnotes are numbered after page breaks or section breaks. Read/write WdNumberingRule.
