# MailMergeDataFields

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002091A-0000-0000-C000-000000000046}  

A collection of MailMergeDataField objects that represent the data fields in a mail merge data source.

**Remarks:** Use the DataFields property to return the MailMergeDataFields collection. The following example displays the names of all the fields in the attached data source. You cannot add fields to the MailMergeDataFields collection. When a data field is added to a data source, the field is automatically included in the MailMergeDataFields collection. Use the EditDataSource method to edit the contents of a data source. The following example adds a data field named "Author" to a table in the attached data source. Use DataFields (Index), where Index is the data field name or the index number, to return a single MailMergeDataField object. The index number represents the position of the data field in the mail merge data source. The following example retrieves the first value from the FName field in the data source attached to the active document. The following example displays the name of first data field in the data source attached to the active document.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified MailMergeDataFields object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of mail merge data fields in the collection. Read-only.

## Methods (1)

- `Item(Index As Variant) As MailMergeDataField`  
  Returns an individual MailMergeDataField object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
