# LineFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209CA-0000-0000-C000-000000000046}  

Represents line and arrowhead formatting. For a line, the LineFormat object contains formatting information for the line itself; for a shape with a border, this object contains formatting information for the shape's border.

**Remarks:** Use the Line property to return a LineFormat object. The following example adds a blue, dashed line to the active document. There is a short, narrow oval at the line's starting point and a long, wide triangle at its endpoint.

## Properties (18)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified LineFormat object.
- `BackColor As ColorFormat  (read-only)`  
  Returns or sets a ColorFormat object that represents the background color for a patterned line. Read/write.
- `BeginArrowheadLength As MsoArrowheadLength  (read/write)`  
  Returns or sets the length of the arrowhead at the beginning of the specified line. Read/write MsoArrowheadLength.
- `BeginArrowheadStyle As MsoArrowheadStyle  (read/write)`  
  Returns or sets the style of the arrowhead at the beginning of the specified line. Read/write MsoArrowheadStyle.
- `BeginArrowheadWidth As MsoArrowheadWidth  (read/write)`  
  Returns or sets the width of the arrowhead at the beginning of the specified line. Read/write MsoArrowheadWidth.
- `DashStyle As MsoLineDashStyle  (read/write)`  
  Returns or sets the dash style for the specified line. Read/write MsoLineDashStyle.
- `EndArrowheadLength As MsoArrowheadLength  (read/write)`  
  Returns or sets the length of the arrowhead at the end of the specified line. Read/write MsoArrowheadLength.
- `EndArrowheadStyle As MsoArrowheadStyle  (read/write)`  
  Returns or sets the style of the arrowhead at the end of the specified line. Read/write MsoArrowheadStyle.
- `EndArrowheadWidth As MsoArrowheadWidth  (read/write)`  
  Returns or sets the width of the arrowhead at the end of the specified line. Read/write MsoArrowheadWidth.
- `ForeColor As ColorFormat  (read-only)`  
  Returns or sets a ColorFormat object that represents the foreground color for the line. Read/write.
- `Pattern As MsoPatternType  (read/write)`  
  Returns or sets a value that represents the pattern applied to the specified line. Read/write MsoPatternType.
- `Style As MsoLineStyle  (read/write)`  
  Returns or sets the line format style. Read/write MsoLineStyle.
- `Transparency As Single  (read/write)`  
  Returns or sets the degree of transparency of line. Read/write Single.
- `Visible As MsoTriState  (read/write)`  
  True if the specified object, or the formatting applied to it, is visible. Read/write MsoTriState.
- `Weight As Single  (read/write)`  
  Returns or sets the thickness of the specified line in points. Read/write Single.
- `InsetPen As MsoTriState  (read/write)`  
  MsoTrue to draw lines inside a specified shape. Read/write MsoTriState.
