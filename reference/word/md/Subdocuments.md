# Subdocuments

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020988-0000-0000-C000-000000000046}  

A collection of Subdocument objects that represent the subdocuments in a range or document.

**Remarks:** Use the Subdocuments property to return the Subdocuments collection. The following example expands all the subdocuments in the active document. Use the AddFromFile or AddFromRange method to add a subdocument to a document. The following example adds a subdocument named "Setup.doc" at the end of the active document. The following example applies the Heading 1 style to the first paragraph in the selection and then creates a subdocument for the contents of the selection. Use Subdocuments (Index), where Index is the index number, to return a single Subdocument object. The following example displays the path and file name of the first subdocument in the active document.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Subdocuments object.
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of subdocuments in the collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`
- `Expanded As Boolean  (read/write)`  
  True if the subdocuments in the specified document are expanded. Read/write Boolean.

## Methods (6)

- `Item(Index As Long) As Subdocument`  
  Returns an individual Subdocument object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
- `AddFromFile(Name As Variant, [ConfirmConversions As Variant], [ReadOnly As Variant], [PasswordDocument As Variant], [PasswordTemplate As Variant], [Revert As Variant], [WritePasswordDocument As Variant], [WritePasswordTemplate As Variant]) As Subdocument`  
  Adds the specified subdocument to the master document at the start of the selection and returns a Subdocument object.
    - `Name As Variant` (required): The file name of the subdocument to be inserted into the master document.
    - `ConfirmConversions As Variant` (optional): True to confirm file conversion in the Convert File dialog box if the file isn't in Word format.
    - `ReadOnly As Variant` (optional): True to insert the subdocument as a read-only document.
    - `PasswordDocument As Variant` (optional): The password required to open the subdocument if it is password protected.
    - `PasswordTemplate As Variant` (optional): The password required to open the template attached to the subdocument if the template is password protected.
    - `Revert As Variant` (optional): Controls what happens if Name is the file name of an open document. True to insert the saved version of the subdocument. False to insert the open version of the subdocument, which may contain unsaved changes.
    - `WritePasswordDocument As Variant` (optional): The password required to save changes to the document file if it is write-protected.
    - `WritePasswordTemplate As Variant` (optional): The password required to save changes to the template attached to the subdocument if the template is write-protected.
- `AddFromRange(Range As Range) As Subdocument`  
  Creates one or more subdocuments from the text in the specified range and returns a SubDocument object.
    - `Range As Range` (required): The range used to create one or more subdocuments.
- `Merge([FirstSubdocument As Variant], [LastSubdocument As Variant])`  
  Merges the specified subdocuments of a master document into a single subdocument.
    - `FirstSubdocument As Variant` (optional): The path and file name of the original document you want to merge revisions with.
    - `LastSubdocument As Variant` (optional): The last subdocument in a range of subdocuments to be merged.
- `Delete()`  
  Deletes the collection of subdocuments.
- `Select()`  
  Selects the specified subdocument.
