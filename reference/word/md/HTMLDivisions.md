# HTMLDivisions

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209E8-0000-0000-C000-000000000046}  

A collection of HTMLDivision objects that represents the HTML DIV elements that exist in a web document.

**Remarks:** Use the HTMLDivisions property to return the HTMLDivisions collection. Use the Add method to add an HTML division to a web document. This example adds a new HTML division to the active document, adds text to the division, and formats the borders around the division.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified HTMLDivisions object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of HTML DIV elements in the collection. Read-only.
- `NestingLevel As Long  (read-only)`  
  Returns the nesting level of the specified HTML DIV elements. Read-only Long.

## Methods (2)

- `Add([Range As Variant]) As HTMLDivision`  
  Returns an HTMLDivision object that represents a new HTML division added to a web document.
    - `Range As Variant` (optional): An existing HTML division around which to place the new HTML division.
- `Item(Index As Long) As HTMLDivision`  
  Returns an individual HTMLDivision object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
