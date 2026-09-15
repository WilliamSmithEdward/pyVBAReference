# OMathFrac

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {50209974-BA32-4A03-8FA6-BAC56CC056FD}  

Represents a fraction, consisting of a numerator and denominator separated by a fraction bar. The fraction bar can be horizontal or diagonal, depending on the fraction properties.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OMathFrac object.
- `Num As OMath  (read-only)`  
  Returns an OMath object that represents the numerator for a fraction. Read-only.
- `Den As OMath  (read-only)`  
  Returns an OMath object that represents the denominator for an equation that contains a fraction. Read-only.
- `Type As WdOMathFracType  (read/write)`  
  Returns or sets a WdOMathFracType constant that represents the layout of a fraction, whether it is stacked, skewed, linear, or without a fraction bar. Read/write.
