# OMathFunctions

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {497142A4-16FD-42C6-BC58-15D89345FC21}  

Represents a collection of functions or structures that Microsoft Word supports, such as fractions, integrals, sums, and radicals.

## Properties (5)

- `_NewEnum As IUnknown  (read-only)`
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OMathFunctions object.
- `Count As Long  (read-only)`  
  Returns the number of items in the OMathFunctions collection. Read-only Long.

## Methods (2)

- `Item(Index As Long) As OMathFunction`  
  Returns an OMathFunction object that represents the specified item in the collection.
    - `Index As Long` (required): Specifies the ordinal position of the object within the collection.
- `Add(Range As Range, Type As WdOMathFunctionType, [NumArgs As Variant], [NumCols As Variant]) As OMathFunction`  
  Inserts a new structure, such as a fraction, into an equation at the specified position and returns an OMathFunction object that represents the structure.
    - `Range As Range` (required): The place at which to insert an equation.
    - `Type As WdOMathFunctionType` (required): The type of equation to insert.
    - `NumArgs As Variant` (optional): The number of arguments in the equation.
    - `NumCols As Variant` (optional): The number of columns in the equation.
