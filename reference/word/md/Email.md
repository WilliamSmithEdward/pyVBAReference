# Email

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209DD-0000-0000-C000-000000000046}  

Represents an email message.

**Remarks:** Use the Email property to return the Email object. The Email object and its properties are valid only if the active document is an unsent forward, reply, or new email message. This example displays the name of the style associated with the current email author. The author style name is the same as the value returned by the UserName property.

## Properties (4)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Email object.
- `CurrentEmailAuthor As EmailAuthor  (read-only)`  
  Returns an EmailAuthor object that represents the author of the current email message. Read-only.
