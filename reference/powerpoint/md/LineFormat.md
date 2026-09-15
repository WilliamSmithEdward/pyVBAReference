# LineFormat

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {9149347F-5A91-11CF-8700-00AA0060263B}  

Represents line and arrowhead formatting. For a line, the LineFormat object contains formatting information for the line itself; for a shape with a border, this object contains formatting information for the shape's border.

**Example:**

```vba
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.AddLine(100, 100, 200, 300).Line

    .DashStyle = msoLineDashDotDot

    .ForeColor.RGB = RGB(50, 0, 128)

    .BeginArrowheadLength = msoArrowheadShort

    .BeginArrowheadStyle = msoArrowheadOval

    .BeginArrowheadWidth = msoArrowheadNarrow

    .EndArrowheadLength = msoArrowheadLong

    .EndArrowheadStyle = msoArrowheadTriangle

    .EndArrowheadWidth = msoArrowheadWide

End With
```

## Properties (18)

- `Application As Object  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Creator As Long  (read-only)`  
  Returns a Long that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in Microsoft PowerPoint, this property returns the hexadecimal number 50575054. Read-only.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `BackColor As ColorFormat  (read/write)`  
  Returns or sets a ColorFormat object that represents the background color for the specified fill or patterned line. Read/write.
- `BeginArrowheadLength As MsoArrowheadLength  (read/write)`  
  Returns or sets the length of the arrowhead at the beginning of the specified line. Read/write.
- `BeginArrowheadStyle As MsoArrowheadStyle  (read/write)`  
  Returns or sets the style of the arrowhead at the beginning of the specified line. Read/write.
- `BeginArrowheadWidth As MsoArrowheadWidth  (read/write)`  
  Returns or sets the width of the arrowhead at the beginning of the specified line. Read/write.
- `DashStyle As MsoLineDashStyle  (read/write)`  
  Returns or sets the dash style for the specified line. Read/write.
- `EndArrowheadLength As MsoArrowheadLength  (read/write)`  
  Returns or sets the length of the arrowhead at the end of the specified line. Read/write.
- `EndArrowheadStyle As MsoArrowheadStyle  (read/write)`  
  Returns or sets the style of the arrowhead at the end of the specified line. Read/write.
- `EndArrowheadWidth As MsoArrowheadWidth  (read/write)`  
  Returns or sets the width of the arrowhead at the end of the specified line. Read/write.
- `ForeColor As ColorFormat  (read/write)`  
  Returns or sets a ColorFormat object that represents the foreground color for the fill, line, or shadow. Read/write.
- `Pattern As MsoPatternType  (read/write)`  
  Sets or returns a value that represents the pattern applied to the specified line. Read/write.
- `Style As MsoLineStyle  (read/write)`  
  Returns or sets the line style. Read/write.
- `Transparency As Single  (read/write)`  
  Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0 (opaque) and 1.0 (clear). Read/write.
- `Visible As MsoTriState  (read/write)`  
  Returns or sets the visibility of the specified object or the formatting applied to the specified object. Read/write.
- `Weight As Single  (read/write)`  
  Returns or sets the thickness of the specified line, in points. Read/write.
- `InsetPen As MsoTriState  (read/write)`  
  Determines whether to draw lines on the inside of a specified shape. Read/write.
