# OMathLimUpp

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {FC9086C6-0287-4997-B2E1-816C334A22F8}  

Represents the upper limit mathematical construct, consisting of text on the baseline and reduced-size text immediately above it.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OMathLimUpp object.
- `E As OMath  (read-only)`  
  Returns an OMath object that represents the base of the specified equation object. Read-only.
- `Lim As OMath  (read-only)`  
  Returns an OMath object that represents the limit of the upper limit object. Read-only.

## Methods (1)

- `ToLimLow() As OMathFunction`  
  Converts an equation from the upper limit to the lower limit.
