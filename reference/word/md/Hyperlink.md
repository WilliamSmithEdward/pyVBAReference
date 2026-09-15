# Hyperlink

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002099D-0000-0000-C000-000000000046}  

Represents a hyperlink. The Hyperlink object is a member of the Hyperlinks collection.

**Remarks:** Use the Hyperlink property to return a Hyperlink object associated with a shape (a shape can have only one hyperlink). The following example activates the hyperlink associated with the first shape in the active document. Use Hyperlinks (Index), where Index is the index number, to return a single Hyperlink object from a document, range, or selection. The following example activates the first hyperlink in the selection.

## Properties (14)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Hyperlink object.
- `Name As String  (read-only)`  
  Returns the name of the specified object. Read-only String.
- `Type As MsoHyperlinkType  (read-only)`  
  Returns the hyperlink type. Read-only MsoHyperlinkType.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that is contained within a hyperlink.
- `Shape As Shape  (read-only)`  
  Returns a Shape object for the specified hyperlink or diagram node.
- `ExtraInfoRequired As Boolean  (read-only)`  
  True if extra information is required to resolve the specified hyperlink. Read-only Boolean.
- `Address As String  (read/write)`  
  Returns or sets the address (for example, a file name or URL) of the specified hyperlink. Read/write String.
- `SubAddress As String  (read/write)`  
  Returns or sets a named location in the destination of the specified hyperlink. Read/write String.
- `EmailSubject As String  (read/write)`  
  Returns or sets the text string for the specified hyperlink's subject line. Read/write String.
- `ScreenTip As String  (read/write)`  
  Returns or sets the text that appears as a ScreenTip when the mouse pointer is positioned over the specified hyperlink. Read/write String.
- `TextToDisplay As String  (read/write)`  
  Returns or sets the specified hyperlink's visible text in a document. Read/write String.
- `Target As String  (read/write)`  
  Returns or sets the name of the frame or window in which to load the hyperlink. Read/write String.

## Methods (4)

- `Delete()`  
  Deletes the specified hyperlink.
- `Follow([NewWindow As Variant], [AddHistory As Variant], [ExtraInfo As Variant], [Method As Variant], [HeaderInfo As Variant])`  
  Displays a cached document associated with the specified Hyperlink object, if it has already been downloaded. Otherwise, this method resolves the hyperlink, downloads the target document, and displays the document in the appropriate application.
    - `NewWindow As Variant` (optional): True to display the target document in a new window. The default value is False.
    - `AddHistory As Variant` (optional): This argument is reserved for future use.
    - `ExtraInfo As Variant` (optional): A string or byte array that specifies additional information for HTTP to use to resolve the hyperlink. For example, you can use ExtraInfo to specify the coordinates of an image map, the contents of a form, or a FAT file name. The string is either posted or appended, depending on the value of Method. Use the ExtraInfoRequired property to determine whether extra information is required.
    - `Method As Variant` (optional): Specifies the way additional information for HTTP is handled. Can be any MsoExtraInfoMethod constant.
    - `HeaderInfo As Variant` (optional): A string that specifies header information for the HTTP request. The default value is an empty string. You can combine several header lines into a single string by using the following syntax: "string1" & vbCr & "string2". The specified string is automatically converted into ANSI characters. Note that the HeaderInfo argument may overwrite default HTTP header fields.
- `AddToFavorites()`  
  Creates a shortcut to the document or hyperlink and adds it to the Favorites folder.
- `CreateNewDocument(FileName As String, EditNow As Boolean, Overwrite As Boolean)`  
  Creates a new document linked to the specified hyperlink.
    - `FileName As String` (required): The file name of the specified document.
    - `EditNow As Boolean` (required): True to have the specified document open immediately in its associated editing environment. The default value is True.
    - `Overwrite As Boolean` (required): True to overwrite any existing file of the same name in the same folder. False if any existing file of the same name is preserved and the FileName argument specifies a new file name. The default value is False.
