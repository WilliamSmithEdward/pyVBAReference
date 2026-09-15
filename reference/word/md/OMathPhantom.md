# OMathPhantom

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {DB77D541-85C3-42E8-8649-AFBD7CF87866}  

Represents a phantom object, which has two primary uses: (1) adding the spacing of the phantom base without displaying that base or (2) suppressing part of the glyph from spacing considerations.

**Remarks:** Use the OMathPhantom object to display the spacing of part of an equation without displaying it, or to remove the spacing for part of a glyph from spacing considerations.

## Properties (10)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OMathPhantom object.
- `E As OMath  (read-only)`  
  Returns an OMath object that represents the base of the specified equation object. Read-only.
- `Show As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether the contents of a phantom object are visible. Read/write.
- `ZeroWid As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether the width of a phantom object is ignored in the spacing of the layout. Read/write.
- `ZeroAsc As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether the ascent of the phantom contents is ignored in the spacing of the layout. Read/write.
- `ZeroDesc As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether the descent of the phantom contents is ignored in the spacing of the layout. Read/write.
- `Transp As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether a phantom object is transparent. Read/write.
- `Smash As Boolean  (read/write)`  
  Returns or sets a Boolean that represents that the contents of the phantom are visible but that the height is not taken into account in the spacing of the layout. Read/write.
