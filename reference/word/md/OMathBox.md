# OMathBox

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {842C37FE-C76F-4B2B-9B60-C408CB5E838E}  

Represents an invisible box around an equation or part of an equation to which you can apply properties that affect the mathematical or formatting properties, such as line breaks.

## Properties (7)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OMathBox object.
- `E As OMath  (read-only)`  
  Returns an OMath object that represents the base of the specified equation object. Read-only.
- `OpEmu As Boolean  (read/write)`  
  Returns or sets a Boolean that represents that the box and its contents behave as a single operator and inherit the properties of an operator. Read/write.
- `NoBreak As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether breaks are allowed inside the box object. Read/write.
- `Diff As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether the box acts as the mathematical differential, in which case the box receives the appropriate horizontal spacing for a differential. Read/write.
