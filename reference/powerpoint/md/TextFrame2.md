# TextFrame2

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934F9-5A91-11CF-8700-00AA0060263B}  

Represents the text frame in a Shape or ShapeRange object. Contains the text in the text frame and exposes properties and methods that control the alignment and anchoring of the text frame.

**Remarks:** Use the TextFrame2 property of the Shape and ShapeRange objects to return a TextFrame2 object. Use the HasTextFrame property to determine whether a shape or shape range has a text frame, and use the HasText property to determine whether the text frame contains text.

**Example:**

```vba
Public Sub TextFrame2_Example()



    Set pptSlide = ActivePresentation.Slides(1)

    With pptSlide.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame2

        .TextRange.Text = "Here is some sample text"

        .MarginBottom = 10

        .MarginLeft = 10

        .MarginRight = 10

        .MarginTop = 10

    End With



End Sub
```

## Properties (21)

- `Application As Object  (read-only)`  
  Returns an Application object that represents the creator of the specified TextFrame2 object. Read-only.
- `Creator As Long  (read-only)`  
  Returns a Long that represents the four-character creator code for the application in which the specified object was created. Read-only.
- `Parent As Object  (read-only)`  
  Returns the parent object of the specified test frame. Read-only.
- `MarginBottom As Single  (read/write)`  
  Returns or sets the distance (in points) between the bottom of the text frame and the bottom of the inscribed rectangle of the shape that contains the text. Read/write.
- `MarginLeft As Single  (read/write)`  
  Returns or sets the distance (in points) between the left edge of the text frame and the left edge of the inscribed rectangle of the shape that contains the text. Read/write.
- `MarginRight As Single  (read/write)`  
  Returns or sets the distance (in points) between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text. Read/write.
- `MarginTop As Single  (read/write)`  
  Returns or sets the distance (in points) between the top of the text frame and the top of the inscribed rectangle of the shape that contains the text. Read/write.
- `Orientation As MsoTextOrientation  (read/write)`  
  Returns or sets text orientation. Read/write.
- `HorizontalAnchor As MsoHorizontalAnchor  (read/write)`  
  Returns or sets the horizontal alignment of text in a text frame. Read/write.
- `VerticalAnchor As MsoVerticalAnchor  (read/write)`  
  Returns or sets the vertical alignment of text in a text frame. Read/write.
- `PathFormat As MsoPathFormat  (read/write)`  
  Returns or sets the path type for the specified text frame. Read/write.
- `WarpFormat As MsoWarpFormat  (read/write)`  
  Returns or sets the warp format (how the text is warped) for the specified text frame. Read/write.
- `WordArtFormat As MsoPresetTextEffect  (read/write)`  
  Returns or sets the WordArt type for the specified text frame. Read/write.
- `WordWrap As MsoTriState  (read/write)`  
  Determines whether lines of text break automatically to fit inside the shape. Read/write.
- `AutoSize As MsoAutoSize  (read/write)`  
  Returns or sets a value that indicates whether the size of the specified shape is changed automatically to fit text within its boundaries. Read/write.
- `ThreeD As ThreeDFormat  (read-only)`  
  Returns a ThreeDFormat object that represents the three-dimensional formatting of the parent shape. Read-only.
- `HasText As MsoTriState  (read-only)`  
  Indicates whether the shape that contains the specified text frame has text associated with it. Read-only.
- `TextRange As TextRange2  (read-only)`  
  Returns a TextRange2 object (PowerPoint) object that represents the text in the specified text frame. Read-only.
- `Column As TextColumn2  (read-only)`  
  Returns the Column object that represents the columns of the specified text frame. Read-only.
- `Ruler As Ruler2  (read-only)`  
  Returns a Ruler2 object that represents the ruler for the specified text. Read-only.
- `NoTextRotation As MsoTriState  (read/write)`  
  Indicates whether to rotate text with the TextFrame2 rotation. One of the MsoTriState constants. Read/write.

## Methods (1)

- `DeleteText()`  
  Deletes the text from a text frame and all the associated properties of the text, including font attributes.
