# Variable

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020966-0000-0000-C000-000000000046}  

Represents a variable stored as part of a document. Document variables are used to preserve macro settings in between macro sessions. The Variable object is a member of the Variables collection. The Variables collection includes all the document variables in a document or template.

**Remarks:** Use Variables (Index), where Index is the document variable name or the index number, to return a single Variable object. The following example displays the value of the Temp document variable in the active document. The index number represents the position of the document variable in the Variables collection. The last variable added to the Variables collection is index number 1; the second-to-last variable added to the collection is index number 2, and so on. The following example displays the name of the first document variable in the active document. Use the Add method of the Variables collection to add a variable to a document. The following example adds a document variable named "Temp" with a value of 12 to the active document. If you try to add a document variable with a name that already exists in the Variables collection, an error occurs. To avoid this error, you can enumerate the collection before adding any new variables. If the Blue document variable already exists in the active document, the following example sets its value to 6. If this variable does not already exist, this example adds it to the document and sets it to 6.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Variable object.
- `Name As String  (read-only)`  
  Returns the name of the specified object. Read-only String.
- `Value As String  (read/write)`  
  Returns or sets the value of a document variable. Read/write String.
- `Index As Long  (read-only)`  
  Returns a Long that represents the ordinal position of a variable with in the collection of variables. Read-only.

## Methods (1)

- `Delete()`  
  Deletes the specified variable.
