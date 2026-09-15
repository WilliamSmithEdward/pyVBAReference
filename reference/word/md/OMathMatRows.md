# OMathMatRows

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {1B426348-607D-433C-9216-C5D2BF0EF31F}  

Represents a collection of matrix rows. Use the OMathMatRow object to access individual members of the collection.

## Properties (5)

- `_NewEnum As IUnknown  (read-only)`
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OMathMatRows object.
- `Count As Long  (read-only)`  
  Returns the number of items in the OMathMatRows collection. Read-only Long.

## Methods (2)

- `Item(Index As Long) As OMathMatRow`  
  Returns an OMathMatRow object that represents the specified item in the collection.
    - `Index As Long` (required): Specifies the ordinal position of the object within the collection.
- `Add([BeforeRow As Variant]) As OMathMatRow`  
  Creates an equation row and adds it to a matrix and returns an OMathMatRow object.
    - `BeforeRow As Variant` (optional): An existing row in the matrix before which to place the new row.
