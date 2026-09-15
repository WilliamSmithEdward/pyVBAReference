# Browser

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002092E-0000-0000-C000-000000000046}  

Represents the browser tool used to move the insertion point to objects in a document. This tool is composed of the three buttons at the bottom of the vertical scroll bar.

**Remarks:** Use the Browser property to return the Browser object. The following example moves the insertion point just before the next field in the active document. The following example moves the insertion point to the previous table and selects it.

## Properties (4)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Browser object.
- `Target As WdBrowseTarget  (read/write)`  
  Returns or sets the document item that the Previous and Next methods locate. Read/write WdBrowseTarget.

## Methods (2)

- `Next()`  
  Moves the selection to the next item indicated by the browser target. Use the Target property to change the browser target.
- `Previous()`  
  Moves the selection to the previous item indicated by the browser target. Use the Target property to change the browser target.
