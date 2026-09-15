# PlotArea

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {56AFD330-440C-4F4C-A39C-ED306D084D5F}  

Represents the plot area of a chart.

**Remarks:** This is the area where your chart data is plotted. The plot area on a 2D chart contains the data markers, gridlines, data labels, trendlines, and optional chart items placed in the chart area. The plot area on a 3D chart contains all the above items plus the walls, floor, axes, axis titles, and tick-mark labels in the chart. The plot area is surrounded by the chart area. The chart area on a 2D chart contains the axes, the chart title, the axis titles, and the legend. The chart area on a 3D chart contains the chart title and the legend. For information about formatting the chart area, see the ChartArea object.

**Example:**

```vba
With ActiveDocument.InlineShapes(1)
 If .HasChart Then
 With .Chart
 .ChartArea.Border.LineStyle = xlDash
 .PlotArea.Border.LineStyle = xlDot
 End With
 End If
End With
```

## Properties (14)

- `Name As String  (read-only)`  
  Returns the name of the object. Read-only String.
- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `Height As Double  (read/write)`  
  Returns or sets the height, in points, of the object. Read/write Double.
- `Left As Double  (read/write)`  
  Returns or sets the distance, in points, from the left edge of the object to the left edge of the chart area. Read/write Double.
- `Top As Double  (read/write)`  
  Returns or sets the distance, in points, from the top edge of the object to the top of the first row (on a worksheet) or the top of the chart area (on a chart). Read/write Double.
- `Width As Double  (read/write)`  
  Returns or sets the width, in points, of the object. Read/write Double.
- `InsideLeft As Double  (read/write)`  
  Returns or sets the distance, in points, from the chart edge to the inside left edge of the plot area. Read/write Double.
- `InsideTop As Double  (read/write)`  
  Returns or sets the distance, in points, from the chart edge to the inside top edge of the plot area. Read/write Double.
- `InsideWidth As Double  (read/write)`  
  Returns or sets the inside width, in points, of the plot area. Read/write Double.
- `InsideHeight As Double  (read/write)`  
  Returns or sets the inside height, in points, of the plot area. Read/write Double.
- `Position As XlChartElementPosition  (read/write)`  
  Returns or sets the position of the plot area on the chart. Read/write XlChartElementPosition.
- `Format As ChartFormat  (read-only)`  
  Returns the line, fill, and effect formatting for the object. Read-only ChartFormat.
- `Application As Object  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft Word application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.

## Methods (4)

- `Select() As Variant`  
  Selects the object.
- `ClearFormats() As Variant`  
  Clears the formatting of the object.
- `SetProperty(ID As String, Value As Variant)`
- `GetProperty(ID As String) As Variant`
