# Line

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {AE6CE2F5-B9D3-407D-85A8-0F10C63289A4}  

Represents an individual line in a Rectangle object of type wdTextRectangle. Use the Line object and related methods and properties to programmatically define page layout in a document.

**Remarks:** Use the Item method to return a specific Line object. The following example accesses the first line in the first rectangle in the first page of the active document. Use the LineType property to determine whether the specified line is a text line (wdTextLine) or a table row (wdTableRow). Then use the Range property to access the contents and formatting for the line. The following example creates a reference to the table if the specified line type is wdTableRow.

## Properties (10)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Line object.
- `LineType As WdLineType  (read-only)`  
  Returns a wdLineType constant that indicates whether a line is a text line or a table row.
- `Left As Long  (read-only)`  
  Returns a Long that represents the number of pixels from the left edge of the page to the left edge of a line.
- `Top As Long  (read-only)`  
  Returns a Long that represents the number of pixels from the top of the page to the top of a line.
- `Width As Long  (read-only)`  
  Returns the width, in points, of a line. Read only Long.
- `Height As Long  (read-only)`  
  Returns or sets the height of a line.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that is contained within a line.
- `Rectangles As Rectangles  (read-only)`  
  Returns a Rectangles collection that represents a portion of text or graphics in a page in a document.
