# Series

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {92D41A75-F07E-4CA4-AF6F-BEF486AA4E6F}  

Represents a series in a chart.

**Remarks:** The Series object is a member of the SeriesCollection collection.

**Example:**

```vba
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).Interior.Color = RGB(255, 0, 0)

    End If

End With
```

## Properties (49)

- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `AxisGroup As XlAxisGroup  (read/write)`  
  Returns the type of axis group. Read/write XlAxisGroup.
- `ErrorBars As ErrorBars  (read-only)`  
  Returns the error bars for the series. Read-only ErrorBars.
- `Explosion As Long  (read/write)`  
  Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Read/write Long.
- `Formula As String  (read/write)`  
  Returns or sets the object's formula in A1-style notation and in the language of the macro. Read/write String.
- `FormulaLocal As String  (read/write)`  
  Returns or sets the formula for the object, using A1-style references in the language of the user. Read/write String.
- `FormulaR1C1 As String  (read/write)`  
  Returns or sets the formula for the object, using R1C1-style notation in the language of the macro. Read/write String.
- `FormulaR1C1Local As String  (read/write)`  
  Returns or sets the formula for the object, using R1C1-style notation in the language of the user. Read/write String.
- `HasDataLabels As Boolean  (read/write)`  
  True if the series has data labels. Read/write Boolean.
- `HasErrorBars As Boolean  (read/write)`  
  True if the series has error bars. Read/write Boolean.
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
- `Name As String  (read/write)`  
  Returns or sets the name of the object. Read/write String.
- `PictureType As XlChartPictureType  (read/write)`  
  Returns or sets a value that specifies how pictures are displayed on a column or bar picture chart. Read/write XlChartPictureType.
- `PlotOrder As Long  (read/write)`  
  Returns or sets the plot order for the selected series within the chart group. Read/write Long.
- `Smooth As Boolean  (read/write)`  
  True if curve smoothing is enabled for the line chart or scatter chart. Read/write Boolean.
- `Type As Long  (read/write)`  
  Returns or sets the series type. Read/write Long.
- `ChartType As XlChartType  (read/write)`  
  Returns or sets the chart type. Read/write XlChartType.
- `Values As Variant  (read/write)`  
  Returns or sets a collection of all the values in the series. Read/write Variant.
- `XValues As Variant  (read/write)`  
  Returns or sets an array of x values for a chart series. Read/write Variant.
- `BubbleSizes As Variant  (read/write)`  
  Returns or sets a string that refers to the worksheet cells that contain the x-value, y-value, and size data for the bubble chart. Read/write Variant.
- `BarShape As XlBarShape  (read/write)`  
  Returns or sets the shape used for a single series in a 3D bar or column chart. Read/write XlBarShape.
- `ApplyPictToSides As Boolean  (read/write)`  
  True if a picture is applied to the sides of the point or all points in the series. Read/write Boolean.
- `ApplyPictToFront As Boolean  (read/write)`  
  True if a picture is applied to the front of the point or all points in the series. Read/write Boolean.
- `ApplyPictToEnd As Boolean  (read/write)`  
  True if a picture is applied to the end of the point or all points in the series. Read/write Boolean.
- `Has3DEffect As Boolean  (read/write)`  
  True if the series has a three-dimensional appearance. Read/write Boolean.
- `Shadow As Boolean  (read/write)`  
  Returns or sets a value that indicates whether the object has a shadow. Read/write Boolean.
- `HasLeaderLines As Boolean  (read/write)`  
  True if the series has leader lines. Read/write Boolean.
- `LeaderLines As LeaderLines  (read-only)`  
  Returns the leader lines for the series. Read-only LeaderLines.
- `Format As ChartFormat  (read-only)`  
  Returns the line, fill, and effect formatting for the object. Read-only ChartFormat.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `PictureUnit2 As Double  (read/write)`  
  Returns or sets the unit for each picture on the chart if the PictureType property is set to xlStackScale; otherwise, this property is ignored. Read/write Double.
- `Application As Application  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft PowerPoint application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `PlotColorIndex As Long  (read-only)`  
  Returns an index value that is used internally to associate series formatting with chart elements. Read-only.
- `InvertColor As Long  (read/write)`  
  Returns or sets the fill color for negative data points in a series. Read/write.
- `InvertColorIndex As XlColorIndex  (read/write)`  
  Returns or sets the fill color for negative data points in a series. Read/write.
- `IsFiltered As Boolean  (read/write)`  
  Returns or sets a Boolean that determines whether the specified chart series is filtered out from the chart. Read/write.
- `ParentDataLabelOption As XlParentDataLabelOptions  (read/write)`  
  Specifies the parent data label option (banner, overlapping, or none) for the specified series within the chart group. Read/write XlParentDataLabelOptions enumeration (PowerPoint).
- `QuartileCalculationInclusiveMedian As Boolean  (read/write)`  
  True if the series uses an inclusive median quartile calculation method. Read/write Boolean.
- `ValueSortOrder As XlValueSortOrder  (read/write)`
- `GeoProjectionType As XlGeoProjectionType  (read/write)`
- `GeoMappingLevel As XlGeoMappingLevel  (read/write)`
- `RegionLabelOption As XlRegionLabelOptions  (read/write)`
- `SeriesColorGradientStyle As XlSeriesColorGradientStyle  (read/write)`

## Methods (12)

- `ClearFormats() As Variant`  
  Clears the formatting of the object.
- `Copy() As Variant`  
  If the series has a picture fill, copies the picture to the Clipboard.
- `DataLabels([Index As Variant]) As Object`  
  Returns an object that represents either a single data label (a DataLabel object) or a collection of all the data labels for the series (a DataLabels collection).
    - `Index As Variant` (optional): The number of the data label.
- `Delete() As Variant`  
  Deletes the object.
- `ErrorBar(Direction As XlErrorBarDirection, Include As XlErrorBarInclude, Type As XlErrorBarType, [Amount As Variant], [MinusValues As Variant]) As Variant`  
  Applies error bars to the series.
    - `Direction As XlErrorBarDirection` (required): One of the enumeration values that specifies the error bar direction.
    - `Include As XlErrorBarInclude` (required): One of the enumeration values that specifies the error bar parts to include.
    - `Type As XlErrorBarType` (required): One of the enumeration values that specifies the error bar type.
    - `Amount As Variant` (optional): The error amount. Used for only the positive error amount when Type is xlErrorBarTypeCustom.
    - `MinusValues As Variant` (optional): The negative error amount when Type is xlErrorBarTypeCustom.
- `Paste() As Variant`  
  Pastes a picture from the Clipboard as the marker on the selected series.
- `Points([Index As Variant]) As Object`  
  Returns a collection of all the points in the series.
    - `Index As Variant` (optional): The name or number of the point.
- `Select() As Variant`  
  Selects the object.
- `Trendlines([Index As Variant]) As Object`  
  Returns a collection of all the trendlines for the series.
- `ApplyDataLabels([Type As XlDataLabelsType], [LegendKey As Variant], [AutoText As Variant], [HasLeaderLines As Variant], [ShowSeriesName As Variant], [ShowCategoryName As Variant], [ShowValue As Variant], [ShowPercentage As Variant], [ShowBubbleSize As Variant], [Separator As Variant]) As Variant`  
  Applies data labels to a series.
    - `Type As XlDataLabelsType` (optional): The type of data label to apply.
    - `LegendKey As Variant` (optional): True to show the legend key next to the point. The default is False.
    - `AutoText As Variant` (optional): True if the object automatically generates appropriate text based on content.
    - `HasLeaderLines As Variant` (optional): For the Chart and Series objects, True if the series has leader lines.
    - `ShowSeriesName As Variant` (optional): True to enable the series name for the data label; otherwise, False.
    - `ShowCategoryName As Variant` (optional): True to enable the category name for the data label; otherwise, False.
    - `ShowValue As Variant` (optional): True to enable the value for the data label; otherwise, False.
    - `ShowPercentage As Variant` (optional): True to enable the percentage for the data label; otherwise, False.
    - `ShowBubbleSize As Variant` (optional): True to enable the bubble size for the data label; otherwise, False.
    - `Separator As Variant` (optional): The separator for the data label.
- `GetProperty(Id As String) As Variant`
- `SetProperty(Id As String, Value As Variant)`
