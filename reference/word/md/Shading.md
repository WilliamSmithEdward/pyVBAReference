# Shading

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002093A-0000-0000-C000-000000000046}  

Contains shading attributes for an object.

**Remarks:** Use the Shading property to return the Shading object. The following example applies fine gray shading to the first paragraph in the active document. The following example applies shading with different foreground and background colors to the selection. The following example applies a vertical line texture to the first row in the first table in the active document.

## Properties (8)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Shading object.
- `ForegroundPatternColorIndex As WdColorIndex  (read/write)`  
  Returns or sets the color that's applied to the foreground of the Shading object. This color is applied to the dots and lines in the shading pattern. Read/write WdColorIndex.
- `BackgroundPatternColorIndex As WdColorIndex  (read/write)`  
  Returns or sets the color that's applied to the background of the Shading object. Read/write WdColorIndex.
- `Texture As WdTextureIndex  (read/write)`  
  Returns or sets the shading texture for the specified object. Read/write WdTextureIndex.
- `ForegroundPatternColor As WdColor  (read/write)`  
  Returns or sets the 24-bit color that's applied to the foreground of the Shading object. This color is applied to the dots and lines in the shading pattern. Read/write.
- `BackgroundPatternColor As WdColor  (read/write)`  
  Returns or sets the 24-bit color that's applied to the background of the Shading object. Read/write.
