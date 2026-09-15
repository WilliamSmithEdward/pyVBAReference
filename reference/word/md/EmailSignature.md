# EmailSignature

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209DC-0000-0000-C000-000000000046}  

Contains information about the email signatures used by Microsoft Word when you create and edit email messages and replies.

**Remarks:** Use the EmailSignature property to return the EmailSignature object. This example changes the signatures Word appends to new outgoing email messages and email message replies.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified EmailSignature object.
- `NewMessageSignature As String  (read/write)`  
  Returns or sets the signature that Microsoft Word appends to new email messages. Read/write String.
- `ReplyMessageSignature As String  (read/write)`  
  Returns or sets the signature that Microsoft Word appends to email message replies. Read/write String.
- `EmailSignatureEntries As EmailSignatureEntries  (read-only)`  
  Returns an EmailSignatureEntries object that represents the email signature entries in Microsoft Word. Read-only.
