# Hyperlink

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024431-0000-0000-C000-000000000046}  

Represents a hyperlink.

**Remarks:** The Hyperlink object is a member of the Hyperlinks collection.

**Example:**

```vba
Worksheets(1).Shapes(1).Hyperlink.Follow NewWindow:=True
```

## Properties (12)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Name As String  (read-only)`  
  Returns a String value that represents the name of the object.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the range that the specified hyperlink is attached to.
- `Shape As Shape  (read-only)`  
  Returns a Shape object that represents the shape attached to the specified hyperlink.
- `SubAddress As String  (read/write)`  
  Returns or sets the location within the document associated with the hyperlink. Read/write String.
- `Address As String  (read/write)`  
  Returns or sets a String value that represents the address of the target document.
- `Type As Long  (read-only)`  
  Returns a Long value, containing an MsoHyperlinkType constant, that represents the location of the HTML frame.
- `EmailSubject As String  (read/write)`  
  Returns or sets the text string of the specified hyperlink's email subject line. The subject line is appended to the hyperlink's address. Read/write String.
- `ScreenTip As String  (read/write)`  
  Returns or sets the ScreenTip text for the specified hyperlink. Read/write String.
- `TextToDisplay As String  (read/write)`  
  Returns or sets the text to be displayed for the specified hyperlink. The default value is the address of the hyperlink. Read/write String.

## Methods (4)

- `AddToFavorites()`  
  Adds a shortcut to the workbook or hyperlink to the Favorites folder.
- `Delete()`  
  Deletes the object.
- `Follow([NewWindow As Variant], [AddHistory As Variant], [ExtraInfo As Variant], [Method As Variant], [HeaderInfo As Variant])`  
  Displays a cached document, if it's already been downloaded. Otherwise, this method resolves the hyperlink, downloads the target document, and displays the document in the appropriate application.
    - `NewWindow As Variant` (optional): True to display the target application in a new window. The default value is False.
    - `AddHistory As Variant` (optional): Not used. Reserved for future use.
    - `ExtraInfo As Variant` (optional): A String or byte array that specifies additional information for HTTP to use to resolve the hyperlink. For example, you can use _ExtraInfo_ to specify the coordinates of an image map, the contents of a form, or a FAT file name.
    - `Method As Variant` (optional): Specifies the way _ExtraInfo_ is attached. Can be one of the MsoExtraInfoMethod constants.
    - `HeaderInfo As Variant` (optional): A String that specifies header information for the HTTP request. The default value is an empty string.
- `CreateNewDocument(Filename As String, EditNow As Boolean, Overwrite As Boolean)`  
  Creates a new document linked to the specified hyperlink.
    - `Filename As String` (required): The file name of the specified document.
    - `EditNow As Boolean` (required): True to have the specified document open immediately in its associated editing environment. The default value is True.
    - `Overwrite As Boolean` (required): True to overwrite any existing file of the same name in the same folder. False if any existing file of the same name is preserved and the _FileName_ argument specifies a new file name. The default value is False.
