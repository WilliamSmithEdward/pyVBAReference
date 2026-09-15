# OMathLimLow

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {74DE9576-8E99-4E28-912B-CB30747C60CE}  

Represents the lower limit mathematical construct, consisting of text on the baseline and reduced-size text immediately below it.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OMathLimLow object.
- `E As OMath  (read-only)`  
  Returns an OMath object that represents the base of the specified equation object. Read-only.
- `Lim As OMath  (read-only)`  
  Returns an OMath object that represents the limit of the lower limit object. Read-only.

## Methods (1)

- `ToLimUpp() As OMathFunction`  
  Converts an equation from the lower limit to the upper limit. .
