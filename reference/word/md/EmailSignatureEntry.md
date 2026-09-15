# EmailSignatureEntry

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209E6-0000-0000-C000-000000000046}  

Represents a single email signature entry. The EmailSignatureEntry object is a member of the EmailSignatureEntries collection. The EmailSignatureEntries collection contains all the email signature entries available to Word.

**Remarks:** Use EmailSignatureEntries (Index), where Index is the email signature entry name or item number, to return a single EmailSignatureEntry object. You must match exactly the spelling (but not necessarily the capitalization) of the name. The following example uses the Delete method to delete the signature entry named "Jeff Smith."

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified EmailSignatureEntry object.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
- `Name As String  (read/write)`  
  Returns or sets the name of the specified object. Read/write String.

## Methods (1)

- `Delete()`  
  Deletes the specified email signature.
