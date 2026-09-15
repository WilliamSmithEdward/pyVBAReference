# OMathBreaks

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {E2E0F3A7-204C-40C5-BAA5-290F374FDF5A}  

Represents a collection of OMathBreak objects that represent all the line breaks in an equation.

## Properties (4)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OMathBreaks object.
- `Count As Long  (read-only)`  
  Returns the number of items in the OMathBreaks collection. Read-only Long.

## Methods (2)

- `Item(Index As Long) As OMathBreak`  
  Returns an OMathBreak object that represents the specified item in the collection.
    - `Index As Long` (required): Specifies the ordinal position of the object within the collection.
- `Add(Range As Range) As OMathBreak`  
  Inserts a break into an equation and returns an OMathBreak object that represents the break.
    - `Range As Range` (required): The position at which to insert the break in the equation.
