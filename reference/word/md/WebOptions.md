# WebOptions

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209E4-0000-0000-C000-000000000046}  

Contains document-level attributes used by Microsoft Word when you save a document as a webpage or open a webpage.

**Remarks:** You can return or set attributes either at the application (global) level or at the document level. (Note that attribute values can be different from one document to another, depending on the attribute value at the time the document was saved.) Document-level attribute settings override application-level attribute settings. Application-level attributes are contained in the DefaultWebOptions object. Use the WebOptions property to return the WebOptions object. The following example checks to see whether PNG (Portable Network Graphics) is allowed as an image format and then sets the _strImageFileType_ variable accordingly.

## Properties (15)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified WebOptions object.
- `OptimizeForBrowser As Boolean  (read/write)`  
  True if Word optimizes the specified Web page for the web browser specified by the BrowserLevel property. Read/write Boolean.
- `BrowserLevel As WdBrowserLevel  (read/write)`  
  Returns or sets WdBrowserLevel that represents the level of web browser at which you want to target the specified Web page. Read/write.
- `RelyOnCSS As Boolean  (read/write)`  
  True if cascading style sheets (CSS) are used for font formatting when you view a saved document in a web browser. The default value is True. Read/write Boolean.
- `OrganizeInFolder As Boolean  (read/write)`  
  True if all supporting files, such as background textures and graphics, are organized in a separate folder when you save the specified document as a webpage. False if supporting files are saved in the same folder as the webpage. The default value is True. Read/write Boolean.
- `UseLongFileNames As Boolean  (read/write)`  
  True if long file names are used when you save the document as a webpage. False if long file names are not used and the DOS file name format (8.3) is used. The default value is True. Read/write Boolean.
- `RelyOnVML As Boolean  (read/write)`  
  True if image files are not generated from drawing objects when you save a document as a webpage. False if images are generated. The default value is False. Read/write Boolean.
- `AllowPNG As Boolean  (read/write)`  
  True if PNG (Portable Network Graphics) is allowed as an image format when you save a document as a webpage. False if PNG is not allowed as an output format. The default value is False. Read/write Boolean.
- `ScreenSize As MsoScreenSize  (read/write)`  
  Returns or sets the ideal minimum screen size (width by height, in pixels) that you should use when viewing the saved document in a web browser. Read/write MsoScreenSize.
- `PixelsPerInch As Long  (read/write)`  
  Returns or sets the density (pixels per inch) of graphics images and table cells on a webpage. Read/write Long.
- `Encoding As MsoEncoding  (read/write)`  
  Returns or sets the document encoding (code page or character set) to be used by the web browser when you view the saved document. Read/write MsoEncoding.
- `FolderSuffix As String  (read-only)`  
  Returns the folder suffix that Microsoft Word uses when you save a document as a webpage, use long file names, and choose to save supporting files in a separate folder (that is, if the UseLongFileNames and OrganizeInFolder properties are set to True). Read-only String.
- `TargetBrowser As MsoTargetBrowser  (read/write)`  
  Sets or returns an MsoTargetBrowser constant representing the target browser for documents viewed in a web browser. Read/write.

## Methods (1)

- `UseDefaultFolderSuffix()`  
  Sets the folder suffix for the specified document to the default suffix for the language support you have selected or installed.
