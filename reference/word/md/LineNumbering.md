# LineNumbering

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020972-0000-0000-C000-000000000046}  

Represents line numbers in the left margin or to the left of each newspaper-style column.

**Remarks:** Use the LineNumbering property to return the LineNumbering object. The following example applies line numbering to the text in the first section of the active document. The following example applies line numbering to the pages in the current section.

## Properties (8)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified LineNumbering object.
- `RestartMode As WdNumberingRule  (read/write)`  
  Returns or sets the way line numbering runs - that is, whether it starts over at the beginning of a new page or section or runs continuously. Read/write WdNumberingRule.
- `StartingNumber As Long  (read/write)`  
  Returns or sets the starting line number. Read/write Long.
- `DistanceFromText As Single  (read/write)`  
  Returns or sets the distance (in points) between the right edge of line numbers and the left edge of the document text. Read/write Single.
- `CountBy As Long  (read/write)`  
  Returns or sets the numeric increment for line numbers. Read/write Long.
- `Active As Long  (read/write)`  
  True if line numbering is active for the specified document, section, or sections. Read/write Long.
