# EmailSignatureEntries

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209E5-0000-0000-C000-000000000046}  

A collection of EmailSignatureEntry objects that represents all the email signature entries available to Word.

**Remarks:** Use the EmailSignatureEntries property to return the EmailSignatureEntries collection. Use the Add method of the EmailSignatureEntries object to add an email signature to Word. The following example creates a new email signature entry based on the author's name and a selection in the active document, and then it sets the new signature entry as the default email signature to use for new messages.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified EmailSignatureEntries object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of email signature entries in the collection. Read-only.

## Methods (2)

- `Item(Index As Variant) As EmailSignatureEntry`  
  Returns an individual EmailSignatureEntry object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `Add(Name As String, Range As Range) As EmailSignatureEntry`  
  Returns an EmailSignatureEntry object that represents a new email signature entry.
    - `Name As String` (required): The name of the email entry.
    - `Range As Range` (required): The range in the document that will be added as the signature.
