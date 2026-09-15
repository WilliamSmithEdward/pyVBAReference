# OMathRad

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {67A7EEC5-285D-4024-B071-BD6B33B88547}  

Represents the mathematical radical object, consisting of a radical, a base, and an optional degree.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OMathRad object.
- `Deg As OMath  (read-only)`  
  Returns an OMath object that represents the degree for a radical. Read-only.
- `E As OMath  (read-only)`  
  Returns an OMath object that represents the base of the specified equation object. Read-only.
- `HideDeg As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether to hide the degree for a radical. Read/write.
