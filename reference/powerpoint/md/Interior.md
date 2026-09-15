# Interior

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {92D41A6C-F07E-4CA4-AF6F-BEF486AA4E6F}  

Represents the interior of an object.

## Properties (9)

- `Color As Variant  (read/write)`  
  Returns or sets the primary color of the object. Read/write Variant.
- `ColorIndex As Variant  (read/write)`  
  Returns or sets the color of the interior. Read/write Variant.
- `InvertIfNegative As Variant  (read/write)`  
  True if Microsoft Word inverts the pattern in the object when it corresponds to a negative number. Read/write Variant.
- `Pattern As Variant  (read/write)`  
  Returns or sets a Variant value, containing an XlPattern constant, that represents the interior pattern.
- `PatternColor As Variant  (read/write)`  
  Returns or sets the color of the interior pattern as an RGB value. Read/write Variant.
- `PatternColorIndex As Variant  (read/write)`  
  Returns or sets the color of the interior pattern as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Read/write Long.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `Application As Application  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft PowerPoint application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
