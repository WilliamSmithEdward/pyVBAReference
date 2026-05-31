# WebPageFont

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0913-0000-0000-C000-000000000046}  

Represents the default font used when documents are saved as webpages for a particular character set.

**Remarks:** Use the WebPageFont object to describe the proportional font, proportional font size, fixed-width font, and fixed-width font size for any available character set.

**Example:**

```vba
With myFont
 ProportionalFont = Verdana
 ProportionalFontSize = 14
```

## Properties (6)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the WebPageFont object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the WebPageFont object was created. Read-only.
- `ProportionalFont As String  (read/write)`  
  Sets or gets the proportional font setting in the host application. Read/write.
- `ProportionalFontSize As Single  (read/write)`  
  Sets or gets the proportional font size setting in the host application, in points. Read/write.
- `FixedWidthFont As String  (read/write)`  
  Sets or gets the fixed-width font setting in the host application. Read/write.
- `FixedWidthFontSize As Single  (read/write)`  
  Sets or gets the fixed-width font size setting in the host application, in points. Read/write.
