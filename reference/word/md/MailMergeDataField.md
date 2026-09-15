# MailMergeDataField

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020919-0000-0000-C000-000000000046}  

Represents a single mail merge field in a data source. The MailMergeDataField object is a member of the MailMergeDataFields collection. The MailMergeDataFields collection includes all the data fields in a mail merge data source (for example, Name, Address, and City).

**Remarks:** Use DataFields (Index), where Index is the data field name or the index number, to return a single MailMergeDataField object. The index number represents the position of the data field in the mail merge data source. The following example retrieves the first value from the FName field in the data source attached to the active document. The following example displays the name of first field in the data source attached to the active document. You cannot add fields to the MailMergeDataFields collection. All data fields in a data source are automatically included in the MailMergeDataFields collection.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified MailMergeDataField object.
- `Value As String  (read-only)`  
  Returns the contents of the mail merge data field or mapped data field for the current record. Read-only String.
- `Name As String  (read-only)`  
  Returns name of the specified object. Read-only String.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
