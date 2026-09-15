# WrapFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209C3-0000-0000-C000-000000000046}  

Represents all the properties for wrapping text around a shape or shape range.

**Remarks:** Use the WrapFormat property to return the WrapFormat object. The following example adds an oval to the active document and specifies that document text wrap around the left and right sides of the square that circumscribes the oval. There will be a 0.1-inch margin between the document text and the top, bottom, left side, and right side of the square.

## Properties (10)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified WrapFormat object.
- `Type As WdWrapType  (read/write)`  
  Returns the wrap type for the specified shape. Read/write WdWrapType.
- `Side As WdWrapSideType  (read/write)`  
  Returns or sets a value that indicates whether the document text should wrap on both sides of the specified shape, on either the left or right side only, or on the side of the shape that's farthest from the page margin.Read/write WdWrapSideType.
- `DistanceTop As Single  (read/write)`  
  Returns or sets the distance (in points) between the document text and the top edge of the text-free area surrounding the specified shape. Read/write Single.
- `DistanceBottom As Single  (read/write)`  
  Returns or sets the distance (in points) between the document text and the bottom edge of the text-free area surrounding the specified shape. Read/write Single.
- `DistanceLeft As Single  (read/write)`  
  Returns or sets the distance (in points) between the document text and the left edge of the text-free area surrounding the specified shape. Read/write Single.
- `DistanceRight As Single  (read/write)`  
  Returns or sets the distance (in points) between the document text and the right edge of the text-free area surrounding the specified shape. Read/write Single.
- `AllowOverlap As Long  (read/write)`  
  Returns or sets a value that specifies whether a given shape can overlap other shapes. Read/write Long.
