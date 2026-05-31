# TextFrame2

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000C0398-0000-0000-C000-000000000046}  

Represents the text frame in a Shape, ShapeRange, or ChartFormat object.

**Remarks:** This object contains the text in the text frame as well as the properties and methods that control the alignment and anchoring of the text frame. Use the TextFrame2 property to return a TextFrame2 object.

**Example:**

```vba
Set myDocument = Worksheets(1)
With myDocument.Shapes.AddShape(msoShapeRectangle, _
 0, 0, 250, 140).TextFrame2
 .TextRange.Text = "Here is some test text"
 .MarginBottom = 10
 .MarginLeft = 10
 .MarginRight = 10
 .MarginTop = 10
End With
```

## Properties (21)

- `Application As Object  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `MarginBottom As Single  (read/write)`  
  Returns or sets the distance (in points) between the bottom of the text frame and the bottom of the inscribed rectangle of the shape that contains the text. Read/write Single.
- `MarginLeft As Single  (read/write)`  
  Returns or sets the distance (in points) between the left edge of the text frame and the left edge of the inscribed rectangle of the shape that contains the text. Read/write Single.
- `MarginRight As Single  (read/write)`  
  Returns or sets the distance (in points) between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text. Read/write Single.
- `MarginTop As Single  (read/write)`  
  Returns or sets the distance (in points) between the top of the text frame and the top of the inscribed rectangle of the shape that contains the text. Read/write Single.
- `Orientation As MsoTextOrientation  (read/write)`  
  Returns or sets a value that represents the text frame orientation. Read/write MsoTextOrientation.
- `HorizontalAnchor As MsoHorizontalAnchor  (read/write)`  
  Returns or sets the horizontal anchor type for the specified text. Read/write MsoHorizontalAnchor.
- `VerticalAnchor As MsoVerticalAnchor  (read/write)`  
  Returns or sets the vertical anchor type for the specified text. Read/write MsoVerticalAnchor.
- `PathFormat As MsoPathFormat  (read/write)`  
  Returns or sets the path type for the specified text frame. Read/write MsoPathFormat.
- `WarpFormat As MsoWarpFormat  (read/write)`  
  Returns or sets the warp type for the specified text frame. Read/write MsoWarpFormat.
- `WordArtformat As MsoPresetTextEffect  (read/write)`  
  Returns or sets the Word Art type for the specified text frame. Read/write MsoPresetTextEffect.
- `WordWrap As MsoTriState  (read/write)`  
  Returns or sets text break lines within or past the boundaries of the shape. Read/write MsoTriState.
- `AutoSize As MsoAutoSize  (read/write)`  
  The size of the specified object that changes automatically to fit text within its boundaries. Read/write MsoAutoSize.
- `ThreeD As ThreeDFormat  (read-only)`  
  Returns a ThreeDFormat object that contains 3D-effect formatting properties for the specified text. Read-only.
- `HasText As MsoTriState  (read-only)`  
  Returns whether the specified text frame has text. Read-only MsoTriState.
- `TextRange As TextRange2  (read-only)`  
  Returns the TextRange2 object that represents the text in the object. Read-only.
- `Column As TextColumn2  (read-only)`  
  Returns the TextColumn2 object that represents the columns within the text frame. Read-only.
- `Ruler As Ruler2  (read-only)`  
  Returns a Ruler2 object that represents the ruler for the specified text. Read-only.
- `NoTextRotation As MsoTriState  (read/write)`  
  Returns or sets whether text remains flat when the specified object is rotated. Read/write.

## Methods (1)

- `DeleteText()`  
  Deletes the text from a text frame and all the associated text properties.
