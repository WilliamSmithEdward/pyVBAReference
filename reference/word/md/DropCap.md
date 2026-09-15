# DropCap

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020956-0000-0000-C000-000000000046}  

Represents a dropped capital letter at the beginning of a paragraph. There is no DropCaps collection; each Paragraph object contains only one DropCap object.

**Remarks:** Use the DropCap property to return a DropCap object. The following example sets a dropped capital letter for the first letter in the first paragraph in the active document.

## Properties (7)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified DropCap object.
- `Position As WdDropPosition  (read/write)`  
  Returns or sets the position of a dropped capital letter. Read/write WdDropPosition.
- `FontName As String  (read/write)`  
  Returns or sets a String that represents the name of the font for the dropped capital letter. Read/write.
- `LinesToDrop As Long  (read/write)`  
  Returns or sets the height (in lines) of the specified dropped capital letter. Read/write Long.
- `DistanceFromText As Single  (read/write)`  
  Returns or sets a Single that represents the distance (in points) between the dropped capital letter and the paragraph text. Read/write.

## Methods (2)

- `Clear()`  
  Removes the dropped capital letter formatting.
- `Enable()`  
  Formats the first character in the specified paragraph as a dropped capital letter.
