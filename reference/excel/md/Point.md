# Point

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002086A-0000-0000-C000-000000000046}  

Represents a single point in a series in a chart.

**Remarks:** The Point object is a member of the Points collection. The Points collection contains all the points in one series.

**Example:**

```vba
Worksheets(1).ChartObjects(1).Chart. _
 SeriesCollection(1).Points(3).MarkerStyle = xlDiamond
```

## Properties (28)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `DataLabel As DataLabel  (read-only)`  
  Returns a DataLabel object that represents the data label associated with the point. Read-only.
- `Explosion As Long  (read/write)`  
  Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie). Read/write Long.
- `HasDataLabel As Boolean  (read/write)`  
  True if the point has a data label. Read/write Boolean.
- `InvertIfNegative As Boolean  (read/write)`  
  True if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number. Read/write Boolean.
- `MarkerBackgroundColor As Long  (read/write)`  
  Sets the marker background color as an RGB value or returns the corresponding color index value. The background color is displayed as the Fill color in the application. Applies only to line, scatter, and radar charts. Read/write Long.
- `MarkerBackgroundColorIndex As XlColorIndex  (read/write)`  
  Returns or sets the marker background color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Applies only to line, scatter, and radar charts. Read/write Long.
- `MarkerForegroundColor As Long  (read/write)`  
  Sets the marker foreground color as an RGB value or returns the corresponding color index value. The foreground color is displayed as the Border color in the application. Applies only to line, scatter, and radar charts. Read/write Long.
- `MarkerForegroundColorIndex As XlColorIndex  (read/write)`  
  Returns or sets the marker foreground color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Applies only to line, scatter, and radar charts. Read/write Long.
- `MarkerSize As Long  (read/write)`  
  Returns or sets the data-marker size, in points. Can be a value from 2 through 72. Read/write Long.
- `MarkerStyle As XlMarkerStyle  (read/write)`  
  Returns or sets the marker style for a point or series in a line chart, scatter chart, or radar chart. Read/write XlMarkerStyle.
- `PictureType As XlChartPictureType  (read/write)`  
  Returns or sets an XlChartPictureType value that represents the way pictures are displayed on a column or bar picture chart.
- `ApplyPictToSides As Boolean  (read/write)`  
  True if a picture is applied to the sides of the point or all points in the series. Read/write Boolean.
- `ApplyPictToFront As Boolean  (read/write)`  
  True if a picture is applied to the front of the point or all points in the series. Read/write Boolean.
- `ApplyPictToEnd As Boolean  (read/write)`  
  True if a picture is applied to the end of the point or all points in the series. Read/write Boolean.
- `Shadow As Boolean  (read/write)`  
  Returns or sets a Boolean value that determines if the object has a shadow.
- `SecondaryPlot As Boolean  (read/write)`  
  True if the point is in the secondary section of either a Pie of Pie chart or a Bar of Pie chart. Applies only to points on Pie of Pie charts or Bar of Pie charts. Read/write Boolean.
- `Has3DEffect As Boolean  (read/write)`  
  True if a point has a three-dimensional appearance. Read/write Boolean.
- `PictureUnit2 As Double  (read/write)`  
  Returns or sets the unit for each picture on the chart if the PictureType property is set to xlStackScale (if not, this property is ignored). Read/write Double.
- `Format As ChartFormat  (read-only)`  
  Returns the ChartFormat object. Read-only.
- `Height As Double  (read-only)`  
  Returns the height, in points, of the object. Read-only.
- `Width As Double  (read-only)`  
  Returns the width, in points, of the object. Read-only.
- `Top As Double  (read-only)`  
  Returns a value that represents the distance, in points, from the top edge of the object to the top edge of the chart area. Read-only.
- `Left As Double  (read-only)`  
  Returns a value that represents the distance, in points, from the left edge of the object to the left edge of the chart area. Read-only.
- `Name As String  (read-only)`  
  Returns the object name. Read-only.
- `IsTotal As Boolean  (read/write)`  
  True if the point represents a total. Read/write Boolean.

## Methods (9)

- `ClearFormats() As Variant`  
  Clears the formatting of the object.
- `Copy() As Variant`  
  If the point has a picture fill, this method copies the picture to the Clipboard.
- `Delete() As Variant`  
  Deletes the series the point belongs to.
- `Paste() As Variant`  
  Pastes a picture from the Clipboard as the marker on the selected point.
- `Select() As Variant`  
  Selects the object.
- `ApplyDataLabels([Type As XlDataLabelsType], [LegendKey As Variant], [AutoText As Variant], [HasLeaderLines As Variant], [ShowSeriesName As Variant], [ShowCategoryName As Variant], [ShowValue As Variant], [ShowPercentage As Variant], [ShowBubbleSize As Variant], [Separator As Variant]) As Variant`  
  Applies data labels to a point.
    - `Type As XlDataLabelsType` (optional): The type of data label to apply.
    - `LegendKey As Variant` (optional): True to show the legend key next to the point. The default value is False.
    - `AutoText As Variant` (optional): True if the object automatically generates appropriate text based on content.
    - `HasLeaderLines As Variant` (optional): For the Chart and Series objects, True if the series has leader lines.
    - `ShowSeriesName As Variant` (optional): Pass a Boolean value to enable or disable the series name for the data label.
    - `ShowCategoryName As Variant` (optional): Pass a Boolean value to enable or disable the category name for the data label.
    - `ShowValue As Variant` (optional): Pass a Boolean value to enable or disable the value for the data label.
    - `ShowPercentage As Variant` (optional): Pass a Boolean value to enable or disable the percentage for the data label.
    - `ShowBubbleSize As Variant` (optional): Pass a Boolean value to enable or disable the bubble size for the data label.
    - `Separator As Variant` (optional): The separator for the data label.
- `PieSliceLocation(loc As XlPieSliceLocation, [Index As XlPieSliceIndex]) As Double`  
  Returns the vertical or horizontal position of a point on a chart item, in points, from the top or left edge of the object to the top or left edge of the chart area.
    - `loc As XlPieSliceLocation` (required): Specifies a horizontal or vertical coordinate.
    - `Index As XlPieSliceIndex` (optional): Specifies which pie slice position coordinate to return. The default value is xlOuterCenterPoint.
- `SetProperty(ID As String, Value As Variant)`
- `GetProperty(ID As String) As Variant`
