# MailMergeFieldName

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002091B-0000-0000-C000-000000000046}  

Represents a mail merge field name in a data source. The MailMergeFieldName object is a member of the MailMergeFieldNames collection. The MailMergeFieldNames collection includes all the data field names in a mail merge data source.

**Remarks:** Use FieldNames (Index), where Index is the index number, to return a single MailMergeFieldName object. The index number represents the position of the field in the mail merge data source. The following example retrieves the name of the last field in the data source attached to the active document. You cannot add fields to the MailMergeFieldNames collection. Field names in a data source are automatically included in the MailMergeFieldNames collection.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified MailMergeFieldName object.
- `Name As String  (read-only)`  
  Returns name of the specified object. Read-only String.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
