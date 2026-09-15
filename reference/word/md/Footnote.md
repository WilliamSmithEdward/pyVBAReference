# Footnote

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002093F-0000-0000-C000-000000000046}  

Represents a footnote positioned at the bottom of the page or beneath text. The Footnote object is a member of the Footnotes collection. The Footnotes collection represents the footnotes in a selection, range, or document.

**Remarks:** Use Footnotes (Index), where Index is the index number, to return a single Footnote object. The index number represents the position of the footnote in the selection, range, or document. The following example applies red formatting to the first footnote in the selection. Use the Add method to add a footnote to the Footnotes collection. The following example inserts an automatically numbered footnote immediately after the selection.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Footnote object.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that's contained within the footnote.
- `Reference As Range  (read-only)`  
  Returns a Range object that represents a footnote reference mark.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.

## Methods (1)

- `Delete()`  
  Deletes the specified footnote.
