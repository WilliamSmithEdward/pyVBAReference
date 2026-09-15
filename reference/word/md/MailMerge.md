# MailMerge

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020920-0000-0000-C000-000000000046}  

Represents the mail merge functionality in Word.

**Remarks:** Use the MailMerge property to return the MailMerge object. The MailMerge object is always available regardless of whether the mail merge operation has begun. Use the State property to determine the status of the mail merge operation. The following example executes a mail merge if the active document is a main document with an attached data source. The following example merges the main document with the first three records in the attached data source and then sends the results to the printer.

## Properties (17)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified MailMerge object.
- `MainDocumentType As WdMailMergeMainDocType  (read/write)`  
  Returns or sets the mail merge main document type. Read/write WdMailMergeMainDocType.
- `State As WdMailMergeState  (read-only)`  
  Returns the current state of a mail merge operation. Read-only WdMailMergeState.
- `Destination As WdMailMergeDestination  (read/write)`  
  Returns or sets the destination of the mail merge results. Read/write WdMailMergeDestination.
- `DataSource As MailMergeDataSource  (read-only)`  
  Returns a MailMergeDataSource object that refers to the data source attached to a mail merge main document. Read-only.
- `Fields As MailMergeFields  (read-only)`  
  Returns a read-only MailMergeFields collection that represents all the mail merge fields in the specified document.
- `ViewMailMergeFieldCodes As Long  (read/write)`  
  True if merge field names are displayed in a mail merge main document. False if information from the current record is displayed. Read/write Long.
- `SuppressBlankLines As Boolean  (read/write)`  
  True if blank lines are suppressed when mail merge fields in a mail merge main document are empty. Read/write Boolean.
- `MailAsAttachment As Boolean  (read/write)`  
  True if the merge documents are sent as attachments when the mail merge destination is an email message or a fax. Read/write Boolean.
- `MailAddressFieldName As String  (read/write)`  
  Returns or sets the name of the field that contains email addresses that are used when the mail merge destination is electronic mail. Read/write String.
- `MailSubject As String  (read/write)`  
  Returns or sets the subject line used when the mail merge destination is electronic mail. Read/write String.
- `HighlightMergeFields As Boolean  (read/write)`  
  True to highlight the merge fields in a document. Read/write Boolean.
- `MailFormat As WdMailMergeMailFormat  (read/write)`  
  Returns a WdMailMergeMailFormat constant that represents the format to use when the mail merge destination is an email message. Read/write.
- `ShowSendToCustom As String  (read/write)`  
  Returns or sets a String corresponding to the caption on a custom button on the Complete the merge step (step six) of the Mail Merge Wizard. Read/write.
- `WizardState As Long  (read/write)`  
  Returns or sets a Long indicating the current Mail Merge Wizard step for a document. The WizardState method returns a number that equates to the current Mail Merge Wizard step; a zero (0) means the Mail Merge Wizard is closed. Read/write.

## Methods (10)

- `CreateDataSource([Name As Variant], [PasswordDocument As Variant], [WritePasswordDocument As Variant], [HeaderRecord As Variant], [MSQuery As Variant], [SQLStatement As Variant], [SQLStatement1 As Variant], [Connection As Variant], [LinkToSource As Variant])`  
  Creates a Microsoft Word document that uses a table to store data for a mail merge.
    - `Name As Variant` (optional): The path and file name for the new data source.
    - `PasswordDocument As Variant` (optional): The password required to open the new data source.
    - `WritePasswordDocument As Variant` (optional): The password required to save changes to the data source.
    - `HeaderRecord As Variant` (optional): Field names for the header record. If this argument is omitted, the standard header record is used: "Title, FirstName, LastName, JobTitle, Company, Address1, Address2, City, State, PostalCode, Country, HomePhone, WorkPhone." To separate field names, use the list separator specified in Regional Settings in Control Panel.
    - `MSQuery As Variant` (optional): True to launch Microsoft Query, if it is installed. The Name, PasswordDocument, and HeaderRecord arguments are ignored.
    - `SQLStatement As Variant` (optional): Defines query options for retrieving data.
    - `SQLStatement1 As Variant` (optional): If the query string is longer than 255 characters, SQLStatement specifies the first portion of the string, and SQLStatement1 specifies the second portion.
    - `Connection As Variant` (optional): A range within which the query specified by SQLStatement will be performed.
    - `LinkToSource As Variant` (optional): True to perform the query specified by Connection and SQLStatement each time the main document is opened.
- `CreateHeaderSource(Name As String, [PasswordDocument As Variant], [WritePasswordDocument As Variant], [HeaderRecord As Variant])`  
  Creates a Microsoft Word document that stores a header record that is used instead of the data source header record in a mail merge.
    - `Name As String` (required): The path and file name for the new header source.
    - `PasswordDocument As Variant` (optional): The password required to open the new header source.
    - `WritePasswordDocument As Variant` (optional): The password required to save changes to the new header source.
    - `HeaderRecord As Variant` (optional): A string that specifies the field names for the header record. If this argument is omitted, the standard header record is used: "Title, FirstName, LastName, JobTitle, Company, Address1, Address2, City, State, PostalCode, Country, HomePhone, WorkPhone." To separate field names in Microsoft Windows, use the list separator specified in Regional Settings in Control Panel.
- `Execute([Pause As Variant])`  
  Performs the specified mail merge operation.
    - `Pause As Variant` (optional): True for Microsoft Word pause and display a troubleshooting dialog box if a mail merge error is found. False to report errors in a new document.
- `Check()`  
  Simulates the mail merge operation, pausing to report each error as it occurs.
- `EditDataSource()`  
  Opens or switches to the mail merge data source.
- `EditHeaderSource()`  
  Opens the header source attached to a mail merge main document, or activates the header source if it is already open.
- `EditMainDocument()`  
  Activates the mail merge main document associated with the specified header source or data source document.
- `OpenDataSource(Name As String, [Format As Variant], [ConfirmConversions As Variant], [ReadOnly As Variant], [LinkToSource As Variant], [AddToRecentFiles As Variant], [PasswordDocument As Variant], [PasswordTemplate As Variant], [Revert As Variant], [WritePasswordDocument As Variant], [WritePasswordTemplate As Variant], [Connection As Variant], [SQLStatement As Variant], [SQLStatement1 As Variant], [OpenExclusive As Variant], [SubType As Variant])`  
  Attaches a data source to the specified document, which becomes a main document if it is not one already.
    - `Name As String` (required): The data source file name. You can specify a Microsoft Query (.qry) file instead of specifying a data source, a connection string, and a query string.
    - `Format As Variant` (optional): The file converter used to open the document. Can be one of the WdOpenFormat constants. To specify an external file format, use the OpenFormat property with the FileConverter object to determine the value to use with this argument.
    - `ConfirmConversions As Variant` (optional): True to display the Convert File dialog box if the file is not in Microsoft Word format.
    - `ReadOnly As Variant` (optional): True to open the data source on a read-only basis.
    - `LinkToSource As Variant` (optional): True to perform the query specified by Connection and SQLStatement each time the main document is opened.
    - `AddToRecentFiles As Variant` (optional): True to add the file name to the list of recently used files at the bottom of the File menu.
    - `PasswordDocument As Variant` (optional): The password used to open the data source. (See Remarks below.)
    - `PasswordTemplate As Variant` (optional): The password used to open the template. (See Remarks below.)
    - `Revert As Variant` (optional): Controls what happens if Name is the file name of an open document. True to discard any unsaved changes to the open document and reopen the file; False to activate the open document.
    - `WritePasswordDocument As Variant` (optional): The password used to save changes to the document. (See Remarks below.)
    - `WritePasswordTemplate As Variant` (optional): The password used to save changes to the template. (See Remarks below.)
    - `Connection As Variant` (optional): A range within which the query specified by SQLStatement is to be performed. (See Remarks below.)
    - `SQLStatement As Variant` (optional): Defines query options for retrieving data. (See Remarks below.)
    - `SQLStatement1 As Variant` (optional): If the query string is longer than 255 characters, SQLStatement specifies the first portion of the string, and SQLStatement1 specifies the second portion. (See Remarks below.)
    - `OpenExclusive As Variant` (optional): True to open exclusively.
    - `SubType As Variant` (optional): Can be one of the WdMergeSubType constants.
- `OpenHeaderSource(Name As String, [Format As Variant], [ConfirmConversions As Variant], [ReadOnly As Variant], [AddToRecentFiles As Variant], [PasswordDocument As Variant], [PasswordTemplate As Variant], [Revert As Variant], [WritePasswordDocument As Variant], [WritePasswordTemplate As Variant], [OpenExclusive As Variant])`  
  Attaches a mail merge header source to the specified document.
    - `Name As String` (required): The file name of the header source.
    - `Format As Variant` (optional): The file converter used to open the document. Can be one of the WdOpenFormat constants. To specify an external file format, use the OpenFormat property with a FileConverter object to determine the value to use with this argument.
    - `ConfirmConversions As Variant` (optional): True to display the Convert File dialog box if the file isn't in Microsoft Word format.
    - `ReadOnly As Variant` (optional): True to open the header source on a read-only basis.
    - `AddToRecentFiles As Variant` (optional): True to add the file name to the list of recently used files at the bottom of the File menu.
    - `PasswordDocument As Variant` (optional): The password required to open the header source document. (See Remarks below.)
    - `PasswordTemplate As Variant` (optional): The password required to open the header source template. (See Remarks below.)
    - `Revert As Variant` (optional): Controls what happens if Name is the file name of an open document. True to discard any unsaved changes to the open document and reopen the file; False to activate the open document.
    - `WritePasswordDocument As Variant` (optional): The password required to save changes to the document data source. (See Remarks below.)
    - `WritePasswordTemplate As Variant` (optional): The password required to save changes to the template data source. (See Remarks below.)
    - `OpenExclusive As Variant` (optional): True to open exclusively.
- `ShowWizard(InitialState As Variant, [ShowDocumentStep As Variant], [ShowTemplateStep As Variant], [ShowDataStep As Variant], [ShowWriteStep As Variant], [ShowPreviewStep As Variant], [ShowMergeStep As Variant])`  
  Displays the Mail Merge Wizard in a document.
    - `InitialState As Variant` (required): The number of the Mail Merge Wizard step to display.
    - `ShowDocumentStep As Variant` (optional): True keeps the "Select document type" step in the sequence of mail merge steps. False removes step one.
    - `ShowTemplateStep As Variant` (optional): True keeps the "Select starting document" step in the sequence of mail merge steps. False removes step two.
    - `ShowDataStep As Variant` (optional): True keeps the "Select recipients" step in the sequence of mail merge steps. False removes step three.
    - `ShowWriteStep As Variant` (optional): True keeps the "Write your letter" step in the sequence of mail merge steps. False removes step four.
    - `ShowPreviewStep As Variant` (optional): True keeps the "Preview your letters" step in the sequence of mail merge steps. False removes step five.
    - `ShowMergeStep As Variant` (optional): True keeps the "Complete the merge" step in the sequence of mail merge steps. False removes step six.
