# WebPageFonts

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0914-0000-0000-C000-000000000046}  

A collection of WebPageFont objects that describe the proportional font, proportional font size, fixed-width font, and fixed-width font size used when documents are saved as webpages. You can specify a different set of webpage font properties for each available character set.

**Remarks:** The WebPageFonts collection contains one WebPageFont object for each character set.

**Example:**

```vba
Dim myFont As WebPageFont
Set myFont = _
 Application.DefaultWebOptions.Fonts.Item_
 (msoCharacterSetEnglishWesternEuropeanOtherLatinScript)
```

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the WebPageFonts object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the WebPageFonts object was created. Read-only.
- `Count As Long  (read-only)`  
  Gets a Long indicating the number of items in the WebPageFonts object. Read-only.
- `Item As WebPageFont  (read-only)`  
  Gets a WebPageFont object from the WebPageFonts collection for a particular value of MsoCharacterSet. Read-only.
- `_NewEnum As IUnknown  (read-only)`
