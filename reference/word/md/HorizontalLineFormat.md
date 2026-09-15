# HorizontalLineFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209DE-0000-0000-C000-000000000046}  

Represents horizontal line formatting.

**Remarks:** Use the HorizontalLineFormat property to return a HorizontalLineFormat object. This example sets the alignment for a new horizontal line. This example adds a horizontal line without any 3D shading. This example adds a horizontal line and sets its length to 50% of the window width.

## Properties (7)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified HorizontalLineFormat object.
- `PercentWidth As Single  (read/write)`  
  Returns or sets the length of the specified horizontal line expressed as a percentage of the window width. Read/write Single.
- `NoShade As Boolean  (read/write)`  
  True if Microsoft Word draws the specified horizontal line without 3D shading. Read/write Boolean.
- `Alignment As WdHorizontalLineAlignment  (read/write)`  
  Returns or sets a WdHorizontalLineAlignment constant that represents the alignment for the specified horizontal line. Read/write.
- `WidthType As WdHorizontalLineWidthType  (read/write)`  
  Returns or sets the width type for the specified HorizontalLineFormat object. Read/write WdHorizontalLineWidthType.
