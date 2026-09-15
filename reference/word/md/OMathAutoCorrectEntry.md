# OMathAutoCorrectEntry

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {D8779F01-4869-4403-B334-D60C5F9C9175}  

Represents an individual entry in the OMathAutoCorrectEntry collection.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OMathAutoCorrectEntry object.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
- `Name As String  (read/write)`  
  Returns or sets a String that represents the name of an equation auto correct entry. Read/write.
- `Value As String  (read/write)`  
  Returns or sets a String that represents the contents of an equation auto correct entry. Read/write.

## Methods (1)

- `Delete()`  
  Deletes the specified equation auto correct entry.
