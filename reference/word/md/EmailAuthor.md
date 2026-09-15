# EmailAuthor

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209D7-0000-0000-C000-000000000046}  

Represents the author of an email message.

**Remarks:** Use the CurrentEmailAuthor property to return the EmailAuthor object. The EmailAuthor object and its properties are valid only if the active document is an unsent forward, reply, or new email message. This example returns the style associated with the current author for unsent replies, forwards, or new email messages, and displays the name of the font associated with this style.

## Properties (4)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified EmailAuthor object.
- `Style As Style  (read-only)`  
  Returns a Style object that represents the style associated with the current email author for unsent replies, forwards, or new email messages.
