# OMathBreak

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {65E515D5-F50B-4951-8F38-FA6AC8707387}  

Represents individual line breaks in an equation. Each OMathBreak object is a member of the OMathBreaks collection.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OMathBreak object.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that is contained in the specified object. Read-only.
- `AlignAt As Long  (read/write)`  
  Returns or sets a Long that represents the operator in one line, to which Microsoft Word aligns consecutive lines in an equation. Read/write.

## Methods (1)

- `Delete()`  
  Deletes the specified equation break.
