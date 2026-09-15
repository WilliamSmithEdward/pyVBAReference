# OMathMatCol

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {CAE36175-3818-4C60-BCBF-0645D51EB33B}  

Represents a matrix column. The OMathMatCol is a member of the OMathMatCols collection.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OMathMatCol object.
- `Args As OMathArgs  (read-only)`  
  Returns an OMathArgs object that represents the arguments in a matrix column. Read-only.
- `ColIndex As Long  (read-only)`  
  Returns a Long that represents the ordinal position of a column within the collection of columns in a matrix. Read-only.
- `Align As WdOMathHorizAlignType  (read/write)`  
  Returns or sets a WdOMathHorizAlignType constant that represents the horizontal alignment for arguments in a matrix column. Read/write.

## Methods (1)

- `Delete()`  
  Deletes the specified matrix column.
