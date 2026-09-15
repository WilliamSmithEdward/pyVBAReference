# Lists

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020993-0000-0000-C000-000000000046}  

A collection of List objects that represent all the lists in the specified document.

**Remarks:** Use the Lists property to return the Lists collection. The following example displays the number of items in each list in the active document. Use Lists (Index), where Index is the index number, to return a single List object. The following example applies the first list format (excluding None) on the Numbered tab in the Bullets and Numbering dialog box to the second list in the active document. When you use a For Each loop to enumerate the Lists collection, the lists in a document are returned in reverse order. The following example counts the items for each list in the active document, from the bottom of the document upward. To add a new list to a document, use the ApplyListTemplate method with the ListFormat object for a specified range. You can manipulate the individual List objects within a document, but for more precise control you should work with the ListFormat object.

## Properties (5)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of lists in the collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Lists object.

## Methods (1)

- `Item(Index As Long) As List`  
  Returns an individual List object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
