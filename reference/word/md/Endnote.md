# Endnote

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002093E-0000-0000-C000-000000000046}  

Represents an endnote. The Endnote object is a member of the Endnotes collection, which represents the endnotes in a selection, range, or document.

**Remarks:** Use Endnotes (Index), where Index is the index number, to return a single Endnote object. The index number represents the position of the endnote in the selection, range, or document. The following example applies red formatting to the first endnote in the selection. Use the Add method to add an endnote to the Endnotes collection. The following example adds an endnote immediately after the selection.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Endnote object.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that is contained in the specified object.
- `Reference As Range  (read-only)`  
  Returns a Range object that represents an endnote reference mark.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.

## Methods (1)

- `Delete()`  
  Deletes the specified endnote.
