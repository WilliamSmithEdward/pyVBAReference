# TextFrame

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493484-5A91-11CF-8700-00AA0060263B}  

Represents the text frame in a Shape object. Contains the text in the text frame and the properties and methods that control the alignment and anchoring of the text frame.

**Example:**

```vba
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes _

        .AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame

    .TextRange.Text = "Here is some test text"

    .MarginBottom = 10

    .MarginLeft = 10

    .MarginRight = 10

    .MarginTop = 10

End With
```

## Properties (15)

- `Application As Object  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Creator As Long  (read-only)`  
  Returns a Long that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in Microsoft PowerPoint, this property returns the hexadecimal number 50575054. Read-only.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
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
- `HasText As MsoTriState  (read-only)`  
  Returns whether the specified shape has text associated with it. Read-only.
- `TextRange As TextRange  (read-only)`  
  Returns a TextRange object that represents the text in the specified text frame. Read-only.
- `Ruler As Ruler  (read-only)`  
  Returns a Ruler object that represents the ruler for the specified text. Read-only.
- `HorizontalAnchor As MsoHorizontalAnchor  (read/write)`  
  Returns or sets the horizontal alignment of text in a text frame. Read/write.
- `VerticalAnchor As MsoVerticalAnchor  (read/write)`  
  Returns or sets the vertical alignment of text in a text frame. Read/write.
- `AutoSize As PpAutoSize  (read/write)`  
  Returns or sets a value that indicates whether the size of the specified shape is changed automatically to fit text within its boundaries. Read/write.
- `WordWrap As MsoTriState  (read/write)`  
  Determines whether lines break automatically to fit inside the shape. Read/write.

## Methods (1)

- `DeleteText()`  
  Deletes the text associated with the specified shape.
