# MailMergeFieldNames

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002091C-0000-0000-C000-000000000046}  

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application. Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified MailMergeFieldNames object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of mail merge field names in the collection. Read-only.

## Methods (1)

- `Item(Index As Variant) As MailMergeFieldName`  
  Returns an individual MailMergeFieldNames object in a collection.
    - `Index As Variant` (required): The individual object to be returned.
