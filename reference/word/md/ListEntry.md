# ListEntry

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020923-0000-0000-C000-000000000046}  

Represents an item in a drop-down form field. The ListEntry object is a member of the ListEntries collection. The ListEntries collection includes all the items in a drop-down form field.

**Remarks:** Use ListEntries (Index), where Index is the list entry name or the index number, to return a single ListEntry object. The index number represents the position of the entry in the drop-down form field (the first item is index number 1). The following example deletes the "Blue" entry from the drop-down form field named "Color." The following example displays the first item in the drop-down form field named "Color." Use the Add method to add an item to a drop-down form field. The following example inserts a drop-down form field and then adds "red," "blue," and "green" to the form field.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ListEntry object.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
- `Name As String  (read/write)`  
  Returns or sets the name of the specified object. Read/write String.

## Methods (1)

- `Delete()`  
  Deletes the specified list entry.
