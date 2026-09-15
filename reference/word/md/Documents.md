# Documents

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002096C-0000-0000-C000-000000000046}  

A collection of all the Document objects that are currently open in Word.

**Remarks:** Use the Documents property to return the Documents collection. The following example displays the names of the open documents. Use the Add method to create a new empty document and add it to the Documents collection. The following example creates a new document based on the Normal template. Use the Open method to open a file. The following example opens the document named "Sales.doc." Use Documents (Index), where Index is the document name or index number to return a single Document object. The following instruction closes the document named "Report.doc" without saving changes. The index number represents the position of the document in the Documents collection. The following example activates the first document in the Documents collection. The following example enumerates the Documents collection to determine whether the document named "Report.doc" is open. If this document is contained in the Documents collection, the document is activated; otherwise, it is opened.

## Properties (5)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of documents in the collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Documents object.

## Methods (9)

- `Item(Index As Variant) As Document`  
  Returns an individual Document object in a collection.
    - `Index As Variant` (required): One-based index of the object to be returned (Long), or the name of the object (String).
- `Close([SaveChanges As Variant], [OriginalFormat As Variant], [RouteDocument As Variant])`  
  Closes the specified documents.
    - `SaveChanges As Variant` (optional): Specifies the save action for the document. Can be one of the following WdSaveOptions constants: wdDoNotSaveChanges, wdPromptToSaveChanges, or wdSaveChanges.
    - `OriginalFormat As Variant` (optional): Specifies the save format for the document. Can be one of the following WdOriginalFormat constants: wdOriginalDocumentFormat, wdPromptUser, or wdWordDocument.
    - `RouteDocument As Variant` (optional): True to route the document to the next recipient. If the document doesn't have a routing slip attached, this argument is ignored.
- `Save([NoPrompt As Variant], [OriginalFormat As Variant])`  
  Saves all the documents in the Documents collection.
    - `NoPrompt As Variant` (optional): True to have Word automatically save all documents. False to have Word prompt the user to save each document that has changed since it was last saved.
    - `OriginalFormat As Variant` (optional): Specifies the way the documents are saved. Can be one of the WdOriginalFormat constants.
- `Add([Template As Variant], [NewTemplate As Variant], [DocumentType As Variant], [Visible As Variant]) As Document`  
  Returns a Document object that represents a new, empty document added to the collection of open documents.
    - `Template As Variant` (optional): The name of the template to be used for the new document. If this argument is omitted, the Normal template is used.
    - `NewTemplate As Variant` (optional): True to open the document as a template. The default value is False.
    - `DocumentType As Variant` (optional): Can be one of the following WdNewDocumentType constants: wdNewBlankDocument, wdNewEmailMessage, wdNewFrameset, or wdNewWebPage. The default constant is wdNewBlankDocument.
    - `Visible As Variant` (optional): True to open the document in a visible window. If this value is False, Microsoft Word opens the document but sets the Visible property of the document window to False. The default value is True.
- `CheckOut(FileName As String)`  
  Copies a specified document from a server to a local computer for editing.
    - `FileName As String` (required): The name of the file to check out.
- `CanCheckOut(FileName As String) As Boolean`  
  True if Microsoft Word can check out a specified document from a server. Read/write Boolean.
    - `FileName As String` (required): The server path and name of the document.
- `Open(FileName As Variant, [ConfirmConversions As Variant], [ReadOnly As Variant], [AddToRecentFiles As Variant], [PasswordDocument As Variant], [PasswordTemplate As Variant], [Revert As Variant], [WritePasswordDocument As Variant], [WritePasswordTemplate As Variant], [Format As Variant], [Encoding As Variant], [Visible As Variant], [OpenAndRepair As Variant], [DocumentDirection As Variant], [NoEncodingDialog As Variant], [XMLTransform As Variant]) As Document`  
  Opens the specified document and adds it to the Documents collection. Returns a Document object.
    - `FileName As Variant` (required): The name of the document (paths are accepted).
    - `ConfirmConversions As Variant` (optional): True to display the Convert File dialog box if the file isn't in Microsoft Word format.
    - `ReadOnly As Variant` (optional): True to open the document as read-only. This argument doesn't override the read-only recommended setting on a saved document. For example, if a document has been saved with read-only recommended turned on, setting the ReadOnly argument to False will not cause the file to be opened as read/write.
    - `AddToRecentFiles As Variant` (optional): True to add the file name to the list of recently used files at the bottom of the File menu.
    - `PasswordDocument As Variant` (optional): The password for opening the document.
    - `PasswordTemplate As Variant` (optional): The password for opening the template.
    - `Revert As Variant` (optional): Controls what happens if FileName is the name of an open document. True to discard any unsaved changes to the open document and reopen the file. False to activate the open document.
    - `WritePasswordDocument As Variant` (optional): The password for saving changes to the document.
    - `WritePasswordTemplate As Variant` (optional): The password for saving changes to the template.
    - `Format As Variant` (optional): The file converter to be used to open the document. Can be one of the WdOpenFormat constants. The default value is wdOpenFormatAuto. To specify an external file format, apply the OpenFormat property to a FileConverter object to determine the value to use with this argument.
    - `Encoding As Variant` (optional): The document encoding (code page or character set) to be used by Microsoft Word when you view the saved document. Can be any valid MsoEncoding constant. For the list of valid MsoEncoding constants, see the Object Browser in the Visual Basic Editor. The default value is the system code page.
    - `Visible As Variant` (optional): True if the document is opened in a visible window. The default value is True.
    - `OpenAndRepair As Variant` (optional): True to repair the document to prevent document corruption.
    - `DocumentDirection As Variant` (optional): Indicates the horizontal flow of text in a document. The default value is wdLeftToRight.
    - `NoEncodingDialog As Variant` (optional): True to skip displaying the Encoding dialog box that Word displays if the text encoding cannot be recognized. The default value is False.
- `OpenNoRepairDialog(FileName As Variant, [ConfirmConversions As Variant], [ReadOnly As Variant], [AddToRecentFiles As Variant], [PasswordDocument As Variant], [PasswordTemplate As Variant], [Revert As Variant], [WritePasswordDocument As Variant], [WritePasswordTemplate As Variant], [Format As Variant], [Encoding As Variant], [Visible As Variant], [OpenAndRepair As Variant], [DocumentDirection As Variant], [NoEncodingDialog As Variant], [XMLTransform As Variant]) As Document`  
  Opens the specified document and adds it to the Documents collection.
    - `FileName As Variant` (required): The name of the document (paths are accepted).
    - `ConfirmConversions As Variant` (optional): True to display the Convert File dialog box if the file is not in Microsoft Word format.
    - `ReadOnly As Variant` (optional): True to open the document as read-only. This argument does not override the read-only recommended setting on a saved document. For example, if a document has been saved with read-only recommended turned on, setting the ReadOnly argument to False will not cause the file to be opened as read/write.
    - `AddToRecentFiles As Variant` (optional): True to add the file name to the list of recently used files at the bottom of the File menu.
    - `PasswordDocument As Variant` (optional): The password for opening the document.
    - `PasswordTemplate As Variant` (optional): The password for opening the template.
    - `Revert As Variant` (optional): Controls what happens if FileName is the name of an open document. True to discard any unsaved changes to the open document and reopen the file. False to activate the open document.
    - `WritePasswordDocument As Variant` (optional): The password for saving changes to the document.
    - `WritePasswordTemplate As Variant` (optional): The password for saving changes to the template.
    - `Format As Variant` (optional): The file converter to be used to open the document. Can be one of the WdOpenFormat constants. The default is wdOpenFormatAuto.
    - `Encoding As Variant` (optional): The document encoding (code page or character set) to be used by Word when you view the saved document. Can be any valid MsoEncoding constant. For the list of valid MsoEncoding constants, see the Object Browser in the Visual Basic Editor. The default is the system code page.
    - `Visible As Variant` (optional): True if the document is opened in a visible window. The default is True.
    - `OpenAndRepair As Variant` (optional): True to repair the document to prevent document corruption.
    - `DocumentDirection As Variant` (optional): Indicates the horizontal flow of text in a document. Can be any valid WdDocumentDirection constant. The default is wdLeftToRight.
    - `NoEncodingDialog As Variant` (optional): True to skip displaying the Encoding dialog box that Word displays if the text encoding cannot be recognized. The default is False.
    - `XMLTransform As Variant` (optional): Specifies a transform to use.
- `AddBlogDocument(ProviderID As String, PostURL As String, BlogName As String, [PostID As String]) As Document`  
  Returns a Document object that represents a new blog document that Microsoft Word publishes to the account described by the first three parameters.
    - `ProviderID As String` (required): A GUID that is the unique value a provider uses when they register themselves with Word.
    - `PostURL As String` (required): The URL that is used to add posts to the blog.
    - `BlogName As String` (required): A display name for the blog that will be used in Word.
    - `PostID As String` (optional): The ID for an existing post with which to populate the document created by using the AddBlogDocument method.
