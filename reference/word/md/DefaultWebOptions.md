# DefaultWebOptions

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209E3-0000-0000-C000-000000000046}  

Contains global application-level attributes used by Microsoft Word when you open a webpage or save a document as a webpage.

**Remarks:** You can return or set attributes either at the application (global) level or at the document level. (Note that attribute values can be different from one document to another, depending on the attribute value at the time the document was saved.) Document-level attribute settings override application-level attribute settings. Document-level attributes are contained in the WebOptions object. Use the DefaultWebOptions method to return the DefaultWebOptions object. The following example checks to see whether PNG (Portable Network Graphics) is allowed as an image format and sets the strImageFileType variable accordingly.

## Properties (21)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified DefaultWebOptions object.
- `OptimizeForBrowser As Boolean  (read/write)`  
  True if Microsoft Word optimizes new Web pages created in Word for the web browser specified by the BrowserLevel property. Read/write Boolean.
- `BrowserLevel As WdBrowserLevel  (read/write)`  
  Returns or sets a WdBrowserLevel constant that represents the level of the web browser for which you want to target new Web pages created in Microsoft Word. Read/write.
- `RelyOnCSS As Boolean  (read/write)`  
  True if cascading style sheets (CSS) are used for font formatting when you view a saved document in a web browser. Read/write Boolean.
- `OrganizeInFolder As Boolean  (read/write)`  
  True if all supporting files, such as background textures and graphics, are organized in a separate folder when you save the specified document as a webpage. False if supporting files are saved in the same folder as the webpage. The default value is True. Read/write Boolean.
- `UpdateLinksOnSave As Boolean  (read/write)`  
  True if hyperlinks and paths to all supporting files are automatically updated before you save the document as a webpage. Read/write Boolean.
- `UseLongFileNames As Boolean  (read/write)`  
  True if long file names are used when you save the document as a webpage. False if long file names are not used and the DOS file name format (8.3) is used. The default value is True. Read/write Boolean.
- `CheckIfOfficeIsHTMLEditor As Boolean  (read/write)`  
  True if Microsoft Word checks to see whether an Office application is the default HTML editor when you start Word. Read/write Boolean.
- `CheckIfWordIsDefaultHTMLEditor As Boolean  (read/write)`  
  True if Microsoft Word checks to see whether it is the default HTML editor when you start Word. Read/write Boolean.
- `RelyOnVML As Boolean  (read/write)`  
  True if image files are not generated from drawing objects when you save a document as a webpage. False if images are generated. The default value is False. Read/write Boolean.
- `AllowPNG As Boolean  (read/write)`  
  False if PNG (Portable Network Graphics) is not allowed as an output format. Read/write Boolean.
- `ScreenSize As MsoScreenSize  (read/write)`  
  Returns or sets the ideal minimum screen size (width by height, in pixels) that you should use when viewing the saved document in a web browser. Read/write MsoScreenSize.
- `PixelsPerInch As Long  (read/write)`  
  Returns or sets the density (pixels per inch) of graphics images and table cells on a webpage. Read/write Long.
- `Encoding As MsoEncoding  (read/write)`  
  Returns or sets the document encoding (code page or character set) to be used by the web browser when you view the saved document. Read/write MsoEncoding.
- `AlwaysSaveInDefaultEncoding As Boolean  (read/write)`  
  True if the default encoding is used when you save a webpage or plain text document, independent of the file's original encoding when opened. Read/write Boolean.
- `Fonts As WebPageFonts  (read-only)`  
  Returns the WebPageFonts collection representing the set of fonts that Microsoft Word uses when you open a webpage in Word.
- `FolderSuffix As String  (read-only)`  
  Returns a String that represents the folder suffix that Microsoft Word uses when you save a document as a webpage, use long file names, or save supporting files in a separate folder. Read-only.
- `TargetBrowser As MsoTargetBrowser  (read/write)`  
  Sets or returns an MsoTargetBrowser constant representing the target browser for documents viewed in a web browser. Read/write.
- `SaveNewWebPagesAsWebArchives As Boolean  (read/write)`  
  True for Microsoft Word to save new Web pages in the Single File Web Page (formerly known as Web Archive) format. Read/write Boolean.
