# OMathBar

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {F08B45F1-8F23-4156-9D63-1820C0ED229A}  

Represents an equation with a bar above or below the base.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OMathBar object.
- `E As OMath  (read-only)`  
  Returns an OMath object that represents the base of the specified equation object. Read-only.
- `BarTop As Boolean  (read/write)`  
  Returns or sets a Boolean that represents the position of a bar in a bar object. True specifies a mathematical overbar. False specifies a mathematical underbar. Read/write.
