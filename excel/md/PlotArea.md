# PlotArea

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208CB-0000-0000-C000-000000000046}  

Represents the plot area of a chart.

**Remarks:** This is the area where your chart data is plotted. The plot area on a 2D chart contains the data markers, gridlines, data labels, trendlines, and optional chart items placed in the chart area. The plot area on a 3D chart contains all the above items plus the walls, floor, axes, axis titles, and tick-mark labels in the chart. The plot area is surrounded by the chart area. The chart area on a 2D chart contains the axes, the chart title, the axis titles, and the legend. The chart area on a 3D chart contains the chart title and the legend. For information about formatting the chart area, see the ChartArea object.

**Example:**

```vba
Charts("Chart1").Activate
With ActiveChart
 .ChartArea.Border.LineStyle = xlDash
 .PlotArea.Border.LineStyle = xlDot
End With
```

## Properties (14)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Name As String  (read-only)`  
  Returns a String value that represents the name of the object.
- `Height As Double  (read/write)`  
  Returns or sets a Double value that represents the height, in points, of the object.
- `Left As Double  (read/write)`  
  Returns or sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).
- `Top As Double  (read/write)`  
  Returns or sets a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).
- `Width As Double  (read/write)`  
  Returns or sets a Double value that represents the width, in points, of the object.
- `InsideLeft As Double  (read/write)`  
  Returns the distance from the chart edge to the inside left edge of the plot area, in points. Read/write Double.
- `InsideTop As Double  (read/write)`  
  Returns the distance from the chart edge to the inside top edge of the plot area, in points. Read/write Double.
- `InsideWidth As Double  (read/write)`  
  Returns the inside width of the plot area, in points. Read/write Double.
- `InsideHeight As Double  (read/write)`  
  Returns the inside height of the plot area, in points. Read/write Double.
- `Position As XlChartElementPosition  (read/write)`  
  Returns or sets the position of the plot area on the chart. Read/write XlChartElementPosition.
- `Format As ChartFormat  (read-only)`  
  Returns the ChartFormat object. Read-only.

## Methods (4)

- `Select() As Variant`  
  Selects the object.
- `ClearFormats() As Variant`  
  Clears the formatting of the object.
- `SetProperty(ID As String, Value As Variant)`
- `GetProperty(ID As String) As Variant`
