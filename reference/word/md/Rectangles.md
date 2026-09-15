# Rectangles

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {7D0F7985-68D9-4D93-91CB-8109280E76CC}  

A collection of Rectangle objects in a page that represent portions of text and graphics. Use the Rectangles collection and related objects and properties for programmatically defining page layout in a document.

**Remarks:** Use the Rectangles property to return a Rectangles collection. The following example returns the Rectangles collection for the first page in the active document.

## Properties (5)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of rectangles in the collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Rectangles object.

## Methods (1)

- `Item(Index As Long) As Rectangle`  
  Returns an individual Rectangle object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
