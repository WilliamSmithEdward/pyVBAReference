# Frame

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002092A-0000-0000-C000-000000000046}  

Represents a frame. The Frame object is a member of the Frames collection. The Frames collection includes all frames in a selection, range, or document.

**Remarks:** Use Frames (Index), where Index is the index number, to return a single Frame object. The index number represents the position of the frame in the selection, range, or document. The following example allows text to wrap around the first frame in the active document. Use the Add method to add a frame around a range. The following example adds a frame around the first paragraph in the active document. You can wrap text around Shape or ShapeRange objects by using the WrapFormat property. You can position a Shape or ShapeRange object by using the Top and Left properties.

## Properties (18)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Frame object.
- `HeightRule As WdFrameSizeRule  (read/write)`  
  Returns or sets a WdFrameSizeRule that represents the rule for determining the height of the specified frame. Read/write.
- `WidthRule As WdFrameSizeRule  (read/write)`  
  Returns or sets the rule used to determine the width of a frame. Read/write WdFrameSizeRule.
- `HorizontalDistanceFromText As Single  (read/write)`  
  Returns or sets the horizontal distance between a frame and the surrounding text, in points. Read/write Single.
- `Height As Single  (read/write)`  
  Returns or sets a Single that represents the height (in points) of the specified frame. Read/write .
- `HorizontalPosition As Single  (read/write)`  
  Returns or sets the horizontal distance between the edge of the frame and the item specified by the RelativeHorizontalPosition property. Read/write Single.
- `LockAnchor As Boolean  (read/write)`  
  True if the specified frame is locked. Read/write Boolean.
- `RelativeHorizontalPosition As WdRelativeHorizontalPosition  (read/write)`  
  Specifies the relative horizontal position of a frame. Read/write WdRelativeHorizontalPosition.
- `RelativeVerticalPosition As WdRelativeVerticalPosition  (read/write)`  
  Specifies the relative vertical position of a frame. Read/write WdRelativeVerticalPosition.
- `VerticalDistanceFromText As Single  (read/write)`  
  Returns or sets the vertical distance (in points) between a frame and the surrounding text. Read/write Single.
- `VerticalPosition As Single  (read/write)`  
  Returns or sets the vertical distance between the edge of the frame and the item specified by the RelativeVerticalPosition property. Read/write Single.
- `Width As Single  (read/write)`  
  Returns or sets the width (in points) of the frame, in points. Read/write Long.
- `TextWrap As Boolean  (read/write)`  
  True if document text wraps around the specified frame. Read/write Boolean.
- `Shading As Shading  (read-only)`  
  Returns a Shading object that refers to the shading formatting for the specified object.
- `Borders As Borders  (read/write)`  
  Returns a Borders collection that represents all the borders for the specified frame.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that's contained within the frame.

## Methods (4)

- `Delete()`  
  Deletes the specified frame.
- `Select()`  
  Selects the specified object.
- `Copy()`  
  Copies the specified frame to the Clipboard.
- `Cut()`  
  Removes the specified frame from the document and places it on the Clipboard.
