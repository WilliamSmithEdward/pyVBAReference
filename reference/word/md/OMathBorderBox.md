# OMathBorderBox

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {2503B6EE-0889-44DF-B920-6D6F9659DEA3}  

Represents an invisible box around an equation or part of an equation to which you can assign properties that affect the layout or mathematical formatting of the entire box. For example, a box can serve as an operator emulator with or without an alignment point, serve as a like breakpoint, have an associated argument size (argSz), or be grouped so as not to allow line breaks.

## Properties (12)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OMathBorderBox object.
- `E As OMath  (read-only)`  
  Returns an OMath object that represents the base of the specified equation object. Read-only.
- `HideTop As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether to hide the top border of an equation's bounding box. Read/write.
- `HideBot As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether to hide the bottom border of an equation's bounding box. Read/write.
- `HideLeft As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether to hide the left border of an equation's bounding box. Read/write.
- `HideRight As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether to hide the right border of an equation's bounding box. Read/write.
- `StrikeH As Boolean  (read/write)`  
  Returns or sets a Boolean that represents a horizontal strikethrough. Read/write.
- `StrikeV As Boolean  (read/write)`  
  Returns or sets a Boolean that represents a vertical strikethrough. Read/write.
- `StrikeBLTR As Boolean  (read/write)`  
  Returns or sets a Boolean that represents a diagonal strikethrough from lower left to upper right. Read/write.
- `StrikeTLBR As Boolean  (read/write)`  
  Returns or sets a Boolean that represents a diagonal strikethrough from upper left to lower right. Read/write.
