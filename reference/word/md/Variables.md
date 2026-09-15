# Variables

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020965-0000-0000-C000-000000000046}  

A collection of Variable objects that represent the variables added to a document or template. Document variables are used to preserve macro settings in between macro sessions.

**Remarks:** Use the Variables property to return the Variables collection. The following example displays the number of variables in the document named "Sales.doc." Use the Add method to add a variable to a document. The following example adds a document variable named "Temp" with a value of 12 to the active document. If you try to add a document variable with a name that already exists in the Variables collection, an error occurs. To avoid this error, you can enumerate the collection before adding any new variables. If the Blue document variable already exists in the active document, the following example sets its value to 6. If this variable doesn't already exist, this example adds it to the document and sets it to 6. Use Variables (Index), where Index is the document variable name or the index number, to return a single Variable object. The following example displays the value of the Temp document variable in the active document. The index number represents the position of the document variable in the Variables collection. The first variable added to the Variables collection is index number 1; the second variable added to the collection is index number 2, and so on.

## Properties (5)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of variables in the collection.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Variables object.

## Methods (2)

- `Item(Index As Variant) As Variable`  
  Returns an individual Variable object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `Add(Name As String, [Value As Variant]) As Variable`  
  Returns a Variable object that represents a variable added to a document.
    - `Name As String` (required): The name of the document variable.
    - `Value As Variant` (optional): The value for the document variable.
