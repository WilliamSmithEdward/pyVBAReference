# OMathScrPre

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {AFAF0C0E-8603-40F6-8FD1-42726CAC21E3}  

Represents an equation that contains a superscript or subscript to the left of the base.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OMathScrPre object.
- `Sub As OMath  (read-only)`  
  Returns an OMath object that represents the subscript for a pre-sub-superscript object. Read-only.
- `Sup As OMath  (read-only)`  
  Returns an OMath object that represents the superscript for a pre-sub-superscript object. Read-only.
- `E As OMath  (read-only)`  
  Returns an OMath object that represents the base of the specified equation object. Read-only.

## Methods (1)

- `ToScrSubSup() As OMathFunction`  
  Converts an equation with a superscript or subscript to the left of the base of the equation to an equation with a base of a superscript or subscript.
