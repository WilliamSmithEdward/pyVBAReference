# OMathMatRow

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {5DAA8BB6-054E-48F6-BEAC-EAAD02BE0CC7}  

Represents a matrix row. The OMathMatRow is a member of the OMathMatRows collection.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OMathMatRow object.
- `Args As OMathArgs  (read-only)`  
  Returns an OMathArgs object that represents the arguments in a matrix row. Read-only.
- `RowIndex As Long  (read-only)`  
  Returns or sets a Long that represents the ordinal position of a matrix row within the collection of matrix rows. Read/write.

## Methods (1)

- `Delete()`  
  Deletes the specified matrix row.
