# ListEntries

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020924-0000-0000-C000-000000000046}  

A collection of ListEntry objects that represent all the items in a drop-down form field.

**Remarks:** Use the ListEntries property to return the ListEntries collection. The following example displays the items that appear in the form field named "Drop1." Use the Add method to add an item to a drop-down form field. The following example inserts a drop-down form field and then adds "red," "blue," and "green" to the form field. Use ListEntries (Index), where Index is the list entry name or the index number, to return a single ListEntry object. The index number represents the position of the entry in the drop-down form field (the first item is index number 1). The following example deletes the "Blue" entry from the drop-down form field named "Color." The following example displays the first item in the drop-down form field named "Color."

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ListEntries object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of list entries in the collection. Read-only.

## Methods (3)

- `Item(Index As Variant) As ListEntry`  
  Returns an individual ListEntry object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `Add(Name As String, [Index As Variant]) As ListEntry`  
  Returns a ListEntry object that represents an item added to a drop-down form field.
    - `Name As String` (required): The name of the drop-down form field item.
    - `Index As Variant` (optional): A number that represents the position of the item in the list.
- `Clear()`  
  Removes all items from a drop-down form field.
