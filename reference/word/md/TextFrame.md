# TextFrame

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209B2-0000-0000-C000-000000000046}  

Represents the text frame in a Shape object. The TextFrame object contains the text in the text frame and the properties that control the margins and orientation of the text frame.

**Remarks:** Use the TextFrame property to return the TextFrame object for a shape. The TextRange property returns a Range object that represents the range of text inside the specified text frame. The following example adds text to the text frame of shape one in the active document. Use the HasText property to determine whether the text frame contains text, as shown in the following example. Text frames can be linked together so that the text flows from the text frame of one shape into the text frame of another shape. Use the Next and Previous properties to link text frames. The following example creates a text box (a rectangle with a text frame) and adds some text to it. It then creates another text box and links the two text frames together so that the text flows from the first text frame into the second one. Use the ContainingRange property to return a Range object that represents the entire story that flows between linked text frames. The following example checks the spelling of the text in TextBox 3 and of any other text that is linked to TextBox 3.

## Properties (23)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Shape  (read-only)`  
  Returns a Shape object that represents the parent shape of the text frame.
- `MarginBottom As Single  (read/write)`  
  Returns or sets the distance (in points) between the bottom of the text frame and the bottom of the inscribed rectangle of the shape that contains the text. Read/write Single.
- `MarginLeft As Single  (read/write)`  
  Returns or sets the distance (in points) between the left edge of the text frame and the left edge of the inscribed rectangle of the shape that contains the text. Read/write Single.
- `MarginRight As Single  (read/write)`  
  Returns or sets the distance (in points) between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text. Read/write Single.
- `MarginTop As Single  (read/write)`  
  Returns or sets the distance (in points) between the top of the text frame and the top of the inscribed rectangle of the shape that contains the text. Read/write Single.
- `Orientation As MsoTextOrientation  (read/write)`  
  Returns or sets the orientation of the text inside the frame. Read/write MsoTextOrientation.
- `TextRange As Range  (read-only)`  
  Returns a Range object that represents the text in the specified text frame.
- `ContainingRange As Range  (read-only)`  
  Returns a Range object that represents the entire story in a series of shapes with linked text frames that the specified text frame belongs to. Read-only.
- `Next As TextFrame  (read/write)`  
  Returns a TextFrame object that represents the next text frame in a collection of shapes. Read-only.
- `Previous As TextFrame  (read/write)`  
  Returns a TextFrame object that represents the previous text frame in a collection of shapes. Read-only.
- `Overflowing As Boolean  (read-only)`  
  True if the text inside the specified text frame doesn't all fit within the frame. Read-only Boolean.
- `HasText As Long  (read-only)`  
  True if the specified shape has text associated with it. Read-only Boolean.
- `AutoSize As Long  (read/write)`  
  Returns or sets a Long that represents whether a text frame is sized automatically. Read/write.
- `WordWrap As Long  (read/write)`  
  True if Microsoft Word wraps Latin text in the middle of a word in the specified text frames. Read/write Long. .
- `VerticalAnchor As MsoVerticalAnchor  (read/write)`  
  Returns or sets an MsoVerticalAnchor constant that represents the vertical alignment of the text within a shape. Read/write.
- `HorizontalAnchor As MsoHorizontalAnchor  (read/write)`  
  Returns or sets the horizontal alignment of text in a text frame. Read/write MsoHorizontalAnchor.
- `PathFormat As MsoPathFormat  (read/write)`  
  Returns or sets the path type for the specified text frame. Read/write MsoPathType.
- `WarpFormat As MsoWarpFormat  (read/write)`  
  Returns or sets the warp format (how the text is warped) for the specified text frame. Read/write MsoWarpFormat.
- `Column As TextColumn2  (read-only)`  
  This object, member, or enumeration is deprecated and is not intended to be used in your code.
- `ThreeD As ThreeDFormat  (read-only)`  
  Returns a ThreeDFormat object that contains 3D effect formatting properties for the specified text frame. Read-only.
- `NoTextRotation As MsoTriState  (read/write)`  
  True if text in the text frame should not rotate when the shape is rotated. Read/write MsoTriState.

## Methods (3)

- `BreakForwardLink()`  
  Breaks the forward link for the specified text frame, if such a link exists.
- `ValidLinkTarget(TargetTextFrame As TextFrame) As Boolean`  
  Determines whether the text frame of one shape can be linked to the text frame of another shape. .
    - `TargetTextFrame As TextFrame` (required): The target text frame to which you want to link the text frame returned by expression.
- `DeleteText()`  
  Deletes the text from a text frame and all the associated properties of the text, including font attributes.
