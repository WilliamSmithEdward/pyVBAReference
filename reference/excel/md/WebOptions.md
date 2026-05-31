# WebOptions

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024449-0000-0000-C000-000000000046}  

Contains workbook-level attributes used by Microsoft Excel when you save a document as a webpage or open a webpage.

**Remarks:** You can return or set attributes either at the application (global) level or at the workbook level. (Note that attribute values can be different from one workbook to another, depending on the attribute value at the time the workbook was saved.) Workbook-level attribute settings override application-level attribute settings. Application-level attributes are contained in the DefaultWebOptions object.

**Example:**

```vba
Set objAppWebOptions = Workbooks(1).WebOptions
With objAppWebOptions
 If .AllowPNG = True Then
 strImageFileType = "PNG"
 Else
 strImageFileType = "JPG"
 End If
End With
```

## Properties (15)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `RelyOnCSS As Boolean  (read/write)`  
  True if cascading style sheets (CSS) are used for font formatting when you view a saved document in a web browser. Microsoft Excel creates a cascading style sheet file and saves it either to the specified folder or to the same folder as your webpage, depending on the value of the OrganizeInFolder property. False if HTML <FONT> tags and cascading style sheets are used. The default value is True. Read/write Boolean.
- `OrganizeInFolder As Boolean  (read/write)`  
  True if all supporting files, such as background textures and graphics, are organized in a separate folder when you save the specified document as a webpage. False if supporting files are saved in the same folder as the webpage. The default value is True. Read/write Boolean.
- `UseLongFileNames As Boolean  (read/write)`  
  True if long file names are used when you save the document as a webpage. False if long file names are not used and the DOS file name format (8.3) is used. The default value is True. Read/write Boolean.
- `DownloadComponents As Boolean  (read/write)`  
  True if the necessary Microsoft Office Web components are downloaded when you view the saved document in a web browser, but only if the components are not already installed. False if the components are not downloaded. The default value is False. Read/write Boolean.
- `RelyOnVML As Boolean  (read/write)`  
  True if image files are not generated from drawing objects when you save a document as a webpage. False if images are generated. The default value is False. Read/write Boolean.
- `AllowPNG As Boolean  (read/write)`  
  True if Portable Network Graphics (PNG) is allowed as an image format when you save documents as a webpage. False if PNG is not allowed as an output format. The default value is False. Read/write Boolean.
- `ScreenSize As MsoScreenSize  (read/write)`  
  Returns or sets the ideal minimum screen size (width by height, in pixels) that you should use when viewing the saved document in a web browser. Can be one of the MsoScreenSize constants. The default constant is msoScreenSize800x600. Read/write MsoScreenSize.
- `PixelsPerInch As Long  (read/write)`  
  Returns or sets the density (pixels per inch) of graphics images and table cells on a webpage. The range of settings is usually from 19 to 480, and common settings for popular screen sizes are 72, 96, and 120. The default setting is 96. Read/write Long.
- `LocationOfComponents As String  (read/write)`  
  Returns or sets the central URL (on the intranet or web) or path (local or network) to the location from which authorized users can download Microsoft Office Web components when viewing your saved document. The default value is the local or network installation path for Microsoft Office. Read/write String.
- `Encoding As MsoEncoding  (read/write)`  
  Returns or sets the document encoding (code page or character set) to be used by the web browser when you view the saved document. The default is the system code page. Read/write MsoEncoding.
- `FolderSuffix As String  (read-only)`  
  Returns the folder suffix that Microsoft Excel uses when you save a document as a webpage, use long file names, and choose to save supporting files in a separate folder (that is, if the UseLongFileNames and OrganizeInFolder properties are set to True). Read-only String.
- `TargetBrowser As MsoTargetBrowser  (read/write)`  
  Returns or sets an MsoTargetBrowser constant indicating the browser version. Read/write.

## Methods (1)

- `UseDefaultFolderSuffix()`  
  Sets the folder suffix for the specified document to the default suffix for the language support that you have selected or installed.
