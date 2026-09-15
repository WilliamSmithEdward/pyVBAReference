# PageNumber

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020987-0000-0000-C000-000000000046}  

Represents a page number in a header or footer. The PageNumber object is a member of the PageNumbers collection. The PageNumbers collection includes all the page numbers in a single header or footer.

**Remarks:** Use PageNumbers (Index), where Index is the index number, to return a single PageNumber object. In most cases, a header or footer will contain only one page number, which is index number 1. The following example centers the first page number in the primary header in section one in the active document. Use the Add method to add a page number (a PAGE field) to a header or footer. The following example adds a page number to the primary footer in the first section and in any subsequent sections. The page number doesn't appear on the first page.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified PageNumber object.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
- `Alignment As WdPageNumberAlignment  (read/write)`  
  Returns or sets a WdPageNumberAlignment constant that represents the alignment for the page number. Read/write.

## Methods (4)

- `Select()`  
  Selects the specified page number.
- `Copy()`  
  Copies the specified page number to the Clipboard.
- `Cut()`  
  Removes the specified object from the document and places it on the Clipboard.
- `Delete()`  
  Deletes the specified page number.
