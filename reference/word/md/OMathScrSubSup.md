# OMathScrSubSup

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {DC489AD4-23C4-4F4B-990F-45A51C7C0C4F}  

Represents an equation with a base that contains a superscript or subscript.

## Properties (7)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OMathScrSubSup object.
- `E As OMath  (read-only)`  
  Returns an OMath object that represents the base of the specified equation object. Read-only.
- `Sub As OMath  (read-only)`  
  Returns an OMath object that represents the subscript for a subscript-superscript object. Read-only.
- `Sup As OMath  (read-only)`  
  Returns an OMath object that represents the superscript for a subscript-superscript object. Read-only.
- `AlignScripts As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether to horizontally align subscripts and superscripts in the sub-superscript object. Read/write.

## Methods (3)

- `RemoveSub() As OMathFunction`  
  Removes the subscript for an equation and returns an OMathFunction object that represents the updated equation without the subscript.
- `RemoveSup() As OMathFunction`  
  Removes the superscript for an equation and returns an OMathFunction object that represents the updated equation without the superscript.
- `ToScrPre() As OMathFunction`  
  Converts an equation with a base superscript or subscript to an equation with a superscript or subscript to the left of the base.
