# OMaths

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {873E774B-926A-4CB1-878D-635A45187595}  

A collection of equations. Use the OMath object to access individual members of the collection.

**Remarks:** Use the Add method to create an equation and add it to a document, selection, or range. The following example creates an equation and uses the BuildUp method of the OMath collection to convert the equation to professional format.

## Properties (5)

- `_NewEnum As IUnknown  (read-only)`
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OMaths object.
- `Count As Long  (read-only)`  
  Returns the number of items in the OMaths collection. Read-only Long.

## Methods (4)

- `Item(Index As Long) As OMath`  
  Returns an OMath object that represents the specified item in the collection.
    - `Index As Long` (required): Specifies the ordinal position of the object within the collection.
- `Linearize()`  
  Converts all equations in the collection to linear format. .
- `BuildUp()`  
  Converts all equations in the collection to professional format.
- `Add(Range As Range) As Range`  
  Creates an equation, from the text equation contained within the specified range, and returns a Range object that contains the new equation.
    - `Range As Range` (required): Specifies a range that contains a text equation.
