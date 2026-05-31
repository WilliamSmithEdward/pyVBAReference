# Series

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002086B-0000-0000-C000-000000000046}  

Represents a series in a chart.

**Remarks:** The Series object is a member of the SeriesCollection collection.

**Example:**

```vba
Worksheets("sheet1").ChartObjects(1).Chart. _
 SeriesCollection(1).Interior.Color = RGB(255, 0, 0)
```

## Properties (52)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `AxisGroup As XlAxisGroup  (read/write)`  
  Returns or sets the group for the specified series. Read/write.
- `ErrorBars As ErrorBars  (read-only)`  
  Returns an ErrorBars object that represents the error bars for the series. Read-only.
- `Explosion As Long  (read/write)`  
  Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie). Read/write Long.
- `Formula As String  (read/write)`  
  Returns or sets a String value that represents the object's formula in A1-style notation and in the language of the macro.
- `FormulaLocal As String  (read/write)`  
  Returns or sets the formula for the object, using A1-style references in the language of the user. Read/write String.
- `FormulaR1C1 As String  (read/write)`  
  Returns or sets the formula for the object, using R1C1-style notation in the language of the macro. Read/write String.
- `FormulaR1C1Local As String  (read/write)`  
  Returns or sets the formula for the object, using R1C1-style notation in the language of the user. Read/write String.
- `HasDataLabels As Boolean  (read/write)`  
  True if the series has data labels. Read/write Boolean.
- `HasErrorBars As Boolean  (read/write)`  
  True if the series has error bars. This property isn't available for 3D charts. Read/write Boolean.
- `InvertIfNegative As Boolean  (read/write)`  
  True if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number. Read/write Boolean.
- `MarkerBackgroundColor As Long  (read/write)`  
  Sets the marker background color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts. Read/write Long.
- `MarkerBackgroundColorIndex As XlColorIndex  (read/write)`  
  Returns or sets the marker background color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Applies only to line, scatter, and radar charts. Read/write Long.
- `MarkerForegroundColor As Long  (read/write)`  
  Sets the marker foreground color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts. Read/write Long.
- `MarkerForegroundColorIndex As XlColorIndex  (read/write)`  
  Returns or sets the marker foreground color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Applies only to line, scatter, and radar charts. Read/write Long.
- `MarkerSize As Long  (read/write)`  
  Returns or sets the data-marker size, in points. Can be a value from 2 through 72. Read/write Long.
- `MarkerStyle As XlMarkerStyle  (read/write)`  
  Returns or sets the marker style for a point or series in a line chart, scatter chart, or radar chart. Read/write XlMarkerStyle.
- `Name As String  (read/write)`  
  Returns or sets a String value representing the name of the object.
- `PictureType As XlChartPictureType  (read/write)`  
  Returns or sets an XlChartPictureType value that represents the way pictures are displayed on a column or bar picture chart.
- `PlotOrder As Long  (read/write)`  
  Returns or sets the plot order for the selected series within the chart group. Read/write Long.
- `Smooth As Boolean  (read/write)`  
  True if curve smoothing is turned on for the line chart or scatter chart. Applies only to line and scatter charts. Read/write.
- `Type As Long  (read/write)`  
  Returns or sets a Long value that represents the series type.
- `ChartType As XlChartType  (read/write)`  
  Returns or sets the chart type. Read/write XlChartType.
- `Values As Variant  (read/write)`  
  Returns or sets a Variant value that represents a collection of all the values in the series.
- `XValues As Variant  (read/write)`  
  Returns or sets an array of x values for a chart series. The XValues property can be set to a range on a worksheet or to an array of values, but it cannot be a combination of both. Read/write Variant.
- `BubbleSizes As Variant  (read/write)`  
  Returns or sets a string that refers to the worksheet cells containing the x-value, y-value, and size data for the bubble chart. When you return the cell reference, it will return a string describing the cells in A1-style notation. To set the size data for the bubble chart, you must use R1C1-style notation. Applies only to bubble charts. Read/write Variant.
- `BarShape As XlBarShape  (read/write)`  
  Returns or sets the shape used with the 3D bar or column chart. Read/write XlBarShape.
- `ApplyPictToSides As Boolean  (read/write)`  
  True if a picture is applied to the sides of the point or all points in the series. Read/write Boolean.
- `ApplyPictToFront As Boolean  (read/write)`  
  True if a picture is applied to the front of the point or all points in the series. Read/write Boolean.
- `ApplyPictToEnd As Boolean  (read/write)`  
  True if a picture is applied to the end of the point or all points in the series. Read/write Boolean.
- `Has3DEffect As Boolean  (read/write)`  
  True if the series has a three-dimensional appearance. Read/write Boolean.
- `Shadow As Boolean  (read/write)`  
  Returns or sets a Boolean value that determines if the object has a shadow.
- `HasLeaderLines As Boolean  (read/write)`  
  True if the series has leader lines. Read/write Boolean.
- `LeaderLines As LeaderLines  (read-only)`  
  Returns a LeaderLines object that represents the leader lines for the series. Read-only.
- `PictureUnit2 As Double  (read/write)`  
  Returns or sets the unit for each picture on the chart if the PictureType property is set to xlStackScale (if not, this property is ignored). Read/write Double.
- `Format As ChartFormat  (read-only)`  
  Returns the ChartFormat object. Read-only.
- `PlotColorIndex As Long  (read-only)`  
  Returns an index value that is used internally to associate series formatting with chart elements. Read-only.
- `InvertColor As Long  (read/write)`  
  Returns or sets the fill color for negative data points in a series. Read/write.
- `InvertColorIndex As Long  (read/write)`  
  Returns or sets the fill color for negative data points in a series. Read/write.
- `IsFiltered As Boolean  (read/write)`  
  This setting controls whether the series has been filtered out from the chart. The default value is False. Read/write Boolean.
- `ParentDataLabelOption As XlParentDataLabelOptions  (read/write)`  
  Specifies the parent data label option (banner, overlapping, or none) for the specified series within the chart group. Read/write XlParentDataLabelOptions.
- `QuartileCalculationInclusiveMedian As Boolean  (read/write)`  
  True if the series uses an inclusive median quartile calculation method. Read/write Boolean.
- `ValueSortOrder As XlValueSortOrder  (read/write)`
- `GeoProjectionType As XlGeoProjectionType  (read/write)`  
  Specifies the geography projection type for the specified series within the chart group. Read/write XlGeoProjectionType.
- `GeoMappingLevel As XlGeoMappingLevel  (read/write)`  
  Specifies the geography mapping level for the specified series within the chart group. Read/write XlGeoMappingLevel.
- `RegionLabelOption As XlRegionLabelOptions  (read/write)`
- `SeriesColorGradientStyle As XlSeriesColorGradientStyle  (read/write)`
- `SeriesColorMinGradientStop As ChartSeriesGradientStopData  (read-only)`
- `SeriesColorMidGradientStop As ChartSeriesGradientStopData  (read-only)`
- `SeriesColorMaxGradientStop As ChartSeriesGradientStopData  (read-only)`

## Methods (12)

- `ClearFormats() As Variant`  
  Clears the formatting of the object.
- `Copy() As Variant`  
  If the series has a picture fill, this method copies the picture to the Clipboard.
- `DataLabels([Index As Variant]) As Object`  
  Returns an object that represents either a single data label (a DataLabel object) or a collection of all the data labels for the series (a DataLabels collection).
    - `Index As Variant` (optional): The number of the data label.
- `Delete() As Variant`  
  Deletes the object.
- `ErrorBar(Direction As XlErrorBarDirection, Include As XlErrorBarInclude, Type As XlErrorBarType, [Amount As Variant], [MinusValues As Variant]) As Variant`  
  Applies error bars to the series. Variant.
    - `Direction As XlErrorBarDirection` (required): The error bar direction.
    - `Include As XlErrorBarInclude` (required): The error bar parts to include.
    - `Type As XlErrorBarType` (required): The error bar type.
    - `Amount As Variant` (optional): The error amount. Used for only the positive error amount when _Type_ is xlErrorBarTypeCustom.
    - `MinusValues As Variant` (optional): The negative error amount when _Type_ is xlErrorBarTypeCustom.
- `Paste() As Variant`  
  Pastes a picture from the Clipboard as the marker on the selected series.
- `Points([Index As Variant]) As Object`  
  Returns an object that represents a single point (a Point object) or a collection of all the points (a Points collection) in the series. Read-only.
    - `Index As Variant` (optional): The name or number of the point.
- `Select() As Variant`  
  Selects the object.
- `Trendlines([Index As Variant]) As Object`  
  Returns an object that represents a single trendline (a Trendline object) or a collection of all the trendlines (a Trendlines collection) for the series.
    - `Index As Variant` (optional): The name or number of the trendline.
- `ApplyDataLabels([Type As XlDataLabelsType], [LegendKey As Variant], [AutoText As Variant], [HasLeaderLines As Variant], [ShowSeriesName As Variant], [ShowCategoryName As Variant], [ShowValue As Variant], [ShowPercentage As Variant], [ShowBubbleSize As Variant], [Separator As Variant]) As Variant`  
  Applies data labels to a series.
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
- `SetProperty(ID As String, Value As Variant)`
- `GetProperty(ID As String) As Variant`
