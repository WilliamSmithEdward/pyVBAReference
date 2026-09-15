# OMathRecognizedFunctions

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {44FEE887-6600-41AB-95A5-DE33C605116C}  

Represents a collection of recognized functions. Use the OMathRecognizedFunction object to access individual members of the collection.

**Remarks:** The OMathRecognizedFunctions collection is a collection of the names that are contained in the Recognized Functions dialog box.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OMathRecognizedFunctions object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns the number of items in the OMathRecognizedFunctions collection. Read-only Long.

## Methods (2)

- `Item(Index As Variant) As OMathRecognizedFunction`  
  Returns an OMathRecognizedFunction object that represents the specified item in the collection.
    - `Index As Variant` (required): Specifies a String or Integer that represents the name or ordinal position of the object in the collection.
- `Add(Name As String) As OMathRecognizedFunction`  
  Creates a new recognized function and returns an OMathRecognizedFunction object.
    - `Name As String` (required): The name of the recognized function.
