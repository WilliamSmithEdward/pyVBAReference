# OMathGroupChar

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {02B17CB4-7D55-4B34-B38B-10381433441F}  

Represents a group character object, consisting of a character drawn above or below text, often with the purpose of visually grouping items.

**Remarks:** Generally, this character is a curly brace, bracket, or arrow but it can be any character.

## Properties (7)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OMathGroupChar object.
- `E As OMath  (read-only)`  
  Returns an OMath object that represents the base of the specified equation object. Read-only.
- `Char As Integer  (read/write)`  
  Returns or sets an Integer that represents the character placed above or below text in a group character object. Read/write.
- `CharTop As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether the grouping character is placed above the base text of the group character object. False displays the group character under the base text. Read/write.
- `AlignTop As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether the grouping character is aligned vertically with the surrounding text or whether the base text that is either above or below the grouping character is aligned vertically with the surrounding text. Read/write.
