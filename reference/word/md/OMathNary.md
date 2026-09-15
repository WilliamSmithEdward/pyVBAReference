# OMathNary

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {CEBD4184-4E6D-4FC6-A42D-2142B1B76AF5}  

Represents the mathematical n-ary object, consisting of an n-ary object, a base (or operand), and optional upper limits and lower limits.

## Properties (11)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OMathNary object.
- `Sub As OMath  (read-only)`  
  Returns an OMath object that represents the lower limit of an n-ary operator. Read-only.
- `Sup As OMath  (read-only)`  
  Returns an OMath object that represents the upper limit of an n-ary operator. Read-only.
- `E As OMath  (read-only)`  
  Returns an OMath object that represents the base of the specified equation object. Read-only.
- `Char As Integer  (read/write)`  
  Returns or sets an Integer that represents a character used as the n-ary operator. Read/write.
- `Grow As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether n-ary operators grow to the full height of the arguments that they contain. Read/write.
- `SubSupLim As Boolean  (read/write)`  
  Returns or sets a Boolean that represents the positioning of n-ary limits in the subscript-superscript or upper limit-lower limit position. Read/write.
- `HideSub As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether to hide the lower limit of an n-ary operator. Read/write.
- `HideSup As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether to hide the upper limit of an n-ary operator. Read/write.
