# MailMergeField

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002091E-0000-0000-C000-000000000046}  

Represents a single mail merge field in a document. The MailMergeDataField object is a member of the MailMergeDataFields collection. The MailMergeDataFields collection includes all the mail merge related fields in a document.

**Remarks:** Use Fields (Index), where Index is the index number, to return a single MailMergeField object. The following example displays the field code of the first mail merge field in the active document. Use the Add method to add a merge field to the MailMergeFields collection. The following example replaces the selection with a MiddleInitial merge field. The MailMergeFields collection has additional methods, such as AddAsk and AddFillIn, for adding fields related to a mail merge operation.

## Properties (8)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified MailMergeField object.
- `Type As WdFieldType  (read-only)`  
  Returns the field type. Read-only WdFieldType.
- `Locked As Boolean  (read/write)`  
  True if the specified field is locked. Read/write Boolean.
- `Code As Range  (read/write)`  
  Returns a Range object that represents a field's code. Read/write.
- `Next As MailMergeField  (read-only)`  
  Returns the next mail merge field in the collection of mail merge fields. Read-only.
- `Previous As MailMergeField  (read-only)`  
  Returns the previous mail merge field in the collection of mail merge fields. Read-only.

## Methods (4)

- `Select()`  
  Selects the specified mail merge field.
- `Copy()`  
  Copies the specified mail merge field to the Clipboard.
- `Cut()`  
  Removes the specified mail merge field from the document and moves it to the Clipboard.
- `Delete()`  
  Deletes the specified mail merge field.
