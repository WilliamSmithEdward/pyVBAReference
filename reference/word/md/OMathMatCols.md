# OMathMatCols

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {EFC71F9C-7F42-4CD4-A7A7-970D7A48CD27}  

Represents a collection of matrix columns. Use the OMathMatCol object to access individual members of the collection.

## Properties (5)

- `_NewEnum As IUnknown  (read-only)`
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OMathMatCols object.
- `Count As Long  (read-only)`  
  Returns the number of items in the OMathMatCols collection. Read-only Long.

## Methods (2)

- `Item(Index As Long) As OMathMatCol`  
  Returns an OMathMatCol object that represents the specified item in the collection.
    - `Index As Long` (required): Specifies the ordinal position of the object within the collection.
- `Add([BeforeCol As Variant]) As OMathMatCol`  
  Creates an equation column and adds it to a matrix and returns an OMathMatCol object.
    - `BeforeCol As Variant` (optional): An existing column in the matrix before which to place the new column.
