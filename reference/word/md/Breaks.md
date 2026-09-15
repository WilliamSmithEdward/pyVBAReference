# Breaks

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {16BE9309-D708-4322-BB1A-B056F58D17EA}  

A collection of page, column, or section breaks in a page. Use the Breaks collection and the related objects and properties to programmatically define page layout in a document.

**Remarks:** Use the Breaks property to return a Breaks collection. The following example returns the breaks in the first page of the active document.

## Properties (5)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns the number of items in the Breaks collection. Read-only Long.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Breaks collection.

## Methods (1)

- `Item(Index As Long) As Break`  
  Returns an individual Break object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
