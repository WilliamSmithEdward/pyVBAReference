# OMathArgs

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {8245795B-9AED-4943-A16D-E586ED8180D1}  

Represents a collection of arguments.

## Properties (4)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OMathArgs object.
- `Count As Long  (read-only)`  
  Returns the number of items in the OMathArgs collection. Read-only Long.

## Methods (2)

- `Item(Index As Long) As OMath`  
  Returns an OMath object that represents the specified item in the collection.
    - `Index As Long` (required): Specifies the ordinal position of the object within the collection.
- `Add([BeforeArg As Variant]) As OMath`  
  Inserts an argument into an equation with variable number of arguments (OMathDelim and OMathEqArray objects) and returns an OMath object.
    - `BeforeArg As Variant` (optional): An existing argument before which to add the new argument.
