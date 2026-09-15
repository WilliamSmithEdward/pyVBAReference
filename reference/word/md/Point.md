# Point

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {8A342FA0-5831-4B5E-82E1-003D0A0C635D}  

Represents a single point in a series in a chart.

**Remarks:** The Point object is a member of the Points collection. The Points collection contains all the points in one series.

**Example:**

```vba
With ActiveDocument.InlineShapes(1)
 If .HasChart Then
 .Chart.SeriesCollection(1).Points(3).MarkerStyle = xlDiamond
 End If
End With
```

## Properties (28)

- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `DataLabel As DataLabel  (read-only)`  
  Returns the data label associated with the point. Read-only DataLabel.
- `Explosion As Long  (read/write)`  
  Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Read/write Long.
- `HasDataLabel As Boolean  (read/write)`  
  True if the point has a data label. Read/write Boolean.
- `InvertIfNegative As Boolean  (read/write)`  
  True if Microsoft Word inverts the pattern in the object when it corresponds to a negative number. Read/write Variant.
- `MarkerBackgroundColor As Long  (read/write)`  
  Sets the marker background color as an RGB value or returns the corresponding color index value. Read/write Long.
- `MarkerBackgroundColorIndex As XlColorIndex  (read/write)`  
  Returns or sets the marker background color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Read/write Long.
- `MarkerForegroundColor As Long  (read/write)`  
  Sets the marker foreground color as an RGB value or returns the corresponding color index value. Read/write Long.
- `MarkerForegroundColorIndex As XlColorIndex  (read/write)`  
  Returns or sets the marker foreground color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Read/write Long.
- `MarkerSize As Long  (read/write)`  
  Returns or sets the data-marker size, in points. Read/write Long.
- `MarkerStyle As XlMarkerStyle  (read/write)`  
  Returns or sets the marker style for a point or series in a line chart, scatter chart, or radar chart. Read/write XlMarkerStyle.
- `PictureType As XlChartPictureType  (read/write)`  
  Returns or sets a value that specifies how pictures are displayed on a column or bar picture chart. Read/write XlChartPictureType.
- `ApplyPictToSides As Boolean  (read/write)`  
  True if a picture is applied to the sides of the point or all points in the series. Read/write Boolean.
- `ApplyPictToFront As Boolean  (read/write)`  
  True if a picture is applied to the front of the point or all points in the series. Read/write Boolean.
- `ApplyPictToEnd As Boolean  (read/write)`  
  True if a picture is applied to the end of the point or all points in the series. Read/write Boolean.
- `Shadow As Boolean  (read/write)`  
  Returns or sets a value that indicates whether the object has a shadow. Read/write Boolean.
- `SecondaryPlot As Boolean  (read/write)`  
  True if the point is in the secondary section of either a pie-of-pie chart or a bar-of-pie chart. Read/write Boolean.
- `Has3DEffect As Boolean  (read/write)`  
  True if a point has a three-dimensional appearance. Read/write Boolean.
- `Format As ChartFormat  (read-only)`  
  Returns the line, fill, and effect formatting for the object. Read-only ChartFormat.
- `Application As Object  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft Word application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `PictureUnit2 As Double  (read/write)`  
  Returns or sets the unit for each picture on the chart if the PictureType property is set to xlStackScale (if not, this property is ignored). Read/write Double.
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
  If the point has a picture fill, copies the picture to the Clipboard.
- `Delete() As Variant`  
  Deletes the object.
- `Paste() As Variant`  
  Pastes a picture from the Clipboard as the marker on the selected point.
- `Select() As Variant`  
  Selects the object.
- `ApplyDataLabels([Type As XlDataLabelsType], [LegendKey As Variant], [AutoText As Variant], [HasLeaderLines As Variant], [ShowSeriesName As Variant], [ShowCategoryName As Variant], [ShowValue As Variant], [ShowPercentage As Variant], [ShowBubbleSize As Variant], [Separator As Variant]) As Variant`  
  Applies data labels to a point.
    - `Type As XlDataLabelsType` (optional): The type of data label to apply. Can be one of the xlDataLabelsType constants.
    - `LegendKey As Variant` (optional): True to show the legend key next to the point. The default is False.
    - `AutoText As Variant` (optional): True if the object automatically generates appropriate text based on content.
    - `HasLeaderLines As Variant` (optional): For the Chart and Series objects, True if the series has leader lines.
    - `ShowSeriesName As Variant` (optional): True to enable the series name for the data label; otherwise, False.
    - `ShowCategoryName As Variant` (optional): True to enable the category name for the data label; otherwise, False.
    - `ShowValue As Variant` (optional): True to enable the value for the data label; otherwise, False.
    - `ShowPercentage As Variant` (optional): True to enable the percentage for the data label; otherwise, False.
    - `ShowBubbleSize As Variant` (optional): True to enable the bubble size for the data label; otherwise, False.
    - `Separator As Variant` (optional): The separator for the data label.
- `PieSliceLocation(loc As XlPieSliceLocation, [Index As XlPieSliceIndex]) As Double`  
  Returns the vertical or horizontal position of a point on a chart item, in points, from the top or left edge of the object to the top or left edge of the chart area.
    - `loc As XlPieSliceLocation` (required): Specifies a horizontal or vertical coordinate.
    - `Index As XlPieSliceIndex` (optional): Specifies which pie slice position coordinate to return. The default value is xlOuterCenterPoint.
- `SetProperty(ID As String, Value As Variant)`
- `GetProperty(ID As String) As Variant`
