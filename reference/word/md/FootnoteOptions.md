# FootnoteOptions

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {BEA85A24-D7DA-4F3D-B58C-ED90FB01D615}  

Represents the properties assigned to a range or selection of footnotes in a document.

**Remarks:** Use the Range or Selection object to return a FootnoteOptions object. Using the FootnoteOptions object, you can assign different footnote properties to different areas of a document. For example, you may want footnotes in the introduction of a long document to be displayed as lowercase letters, while in the rest of your document they are displayed as asterisks. The following example uses the NumberingRule, NumberStyle, and StartingNumber properties to format the footnotes in the first section of the active document.

## Properties (8)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified FootnoteOptions object.
- `Location As WdFootnoteLocation  (read/write)`  
  Returns or sets the position of all footnotes. Read/write WdFootnoteLocation.
- `NumberStyle As WdNoteNumberStyle  (read/write)`  
  Returns or sets the number style for the footnotes. Read/write WdNoteNumberStyle.
- `StartingNumber As Long  (read/write)`  
  Returns or sets the starting footnote number. Read/write Long.
- `NumberingRule As WdNumberingRule  (read/write)`  
  Returns or sets the way footnotes or endnotes are numbered after page breaks or section breaks. Read/write WdNumberingRule.
- `LayoutColumns As Long  (read/write)`  
  Returns or sets the way footnotes are laid out in columns when the section containing the reference mark has multiple columns. Read-Write Long.
