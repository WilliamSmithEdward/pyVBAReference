# Rectangle

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {ADD4EDF3-2F33-4734-9CE6-D476097C5ADA}  

Represents a portion of text or a graphic in a page. Use the Rectangle object and related methods and properties for programmatically defining page layout in a document.

**Remarks:** Use the Item method to return a specific Rectangle object. The following example accesses the first rectangle in the first page of the active document. Use the RectangleType property to determine the type of rectangle. The following example creates a ShapeRange object if the specified rectangle is a shape.

## Properties (10)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Rectangle object.
- `RectangleType As WdRectangleType  (read-only)`  
  Returns a WdRectangleType constant that represents the type for the specified rectangle.
- `Left As Long  (read-only)`  
  Returns a Long that represents the number of pixels from the left edge of the page to the left edge of a rectangle.
- `Top As Long  (read-only)`  
  Returns a Long that represents the number of pixels from the top of the page to the top of a rectangle.
- `Width As Long  (read-only)`  
  Returns or sets a Long that represents the width, in points, of a rectangle. Read/write Long.
- `Height As Long  (read-only)`  
  Returns a Long that represents the height of a rectangle, in pixels.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that is contained within a rectangle.
- `Lines As Lines  (read-only)`  
  Returns a Lines collection that represents the lines in a specified portion of text in a page.
