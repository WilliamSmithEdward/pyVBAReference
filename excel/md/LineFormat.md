# LineFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000C0317-0000-0000-C000-000000000046}  

Represents line and arrowhead formatting.

**Remarks:** For a line, the LineFormat object contains formatting information for the line itself; for a shape with a border, this object contains formatting information for the shape's border.

**Example:**

```vba
Set myDocument = Worksheets(1)
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
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `BackColor As ColorFormat  (read/write)`  
  Returns or sets a ColorFormat object that represents the specified fill background color.
- `BeginArrowheadLength As MsoArrowheadLength  (read/write)`  
  Returns or sets the length of the arrowhead at the beginning of the specified line. Read/write MsoArrowheadLength.
- `BeginArrowheadStyle As MsoArrowheadStyle  (read/write)`  
  Returns or sets the style of the arrowhead at the beginning of the specified line. Read/write MsoArrowheadStyle.
- `BeginArrowheadWidth As MsoArrowheadWidth  (read/write)`  
  Returns or sets the width of the arrowhead at the beginning of the specified line. Read/write MsoArrowheadWidth.
- `DashStyle As MsoLineDashStyle  (read/write)`  
  Returns or sets the dash style for the specified line. Can be one of the MsoLineDashStyle contants. Read/write Long.
- `EndArrowheadLength As MsoArrowheadLength  (read/write)`  
  Returns or sets the length of the arrowhead at the end of the specified line. Read/write MsoArrowheadLength.
- `EndArrowheadStyle As MsoArrowheadStyle  (read/write)`  
  Returns or sets the style of the arrowhead at the end of the specified line. Read/write MsoArrowheadStyle.
- `EndArrowheadWidth As MsoArrowheadWidth  (read/write)`  
  Returns or sets the width of the arrowhead at the end of the specified line. Read/write MsoArrowheadWidth.
- `ForeColor As ColorFormat  (read/write)`  
  Returns or sets a ColorFormat object that represents the specified foreground fill or solid color.
- `Pattern As MsoPatternType  (read/write)`  
  Returns or sets an MsoPatternType value that represents the fill pattern.
- `Style As MsoLineStyle  (read/write)`  
  Returns or sets an MsoLineStyle value that represents the style of the line.
- `Transparency As Single  (read/write)`  
  Returns or sets the degree of transparency of the specified fill as a value from 0.0 (opaque) through 1.0 (clear). Read/write Double.
- `Visible As MsoTriState  (read/write)`  
  Returns or sets an MsoTriState value that determines whether the object is visible. Read/write.
- `Weight As Single  (read/write)`  
  Returns or sets a Single value that represents the weight of the line.
- `InsetPen As MsoTriState  (read/write)`  
  Returns or sets whether lines are drawn inside the specified shape's boundaries. Read/write.
