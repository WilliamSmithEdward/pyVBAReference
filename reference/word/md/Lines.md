# Lines

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {E2E8A400-0615-427D-ADCC-CAD39FFEBD42}  

A collection of Line objects that represents the lines in a Rectangle object that is of type wdTextRectangle.

**Remarks:** Use the Lines property to return a collection of lines for a specified rectangle. The following example accesses the lines in the first rectangle in the first page in the active document. Use the RectangleType property of the specified Rectangle object to determine whether the Rectangle object is of type wdTextRectangle. The following example returns the collection of lines in the first rectangle in the first page of the active document if the specified rectangle contains text.

## Properties (5)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of lines in the collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Lines object.

## Methods (1)

- `Item(Index As Long) As Line`  
  Returns an individual Line object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
