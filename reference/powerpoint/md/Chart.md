# Chart

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {92D41A55-F07E-4CA4-AF6F-BEF486AA4E6F}  

Represents a chart in a presentation.

## Properties (45)

- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `ChartType As XlChartType  (read/write)`  
  Returns or sets the chart type. Read/write XlChartType.
- `HasDataTable As Boolean  (read/write)`  
  True if the chart has a data table. Read/write Boolean.
- `PlotBy As XlRowCol  (read/write)`  
  Returns or sets the way columns or rows are used as data series on the chart. Read/write Long.
- `DataTable As DataTable  (read-only)`  
  Returns the chart data table. Read-only DataTable.
- `BarShape As XlBarShape  (read/write)`  
  Returns or sets the shape used for every series in a 3D bar or column chart. Read/write XlBarShape.
- `SideWall As Walls  (read-only)`  
  Returns a Walls object that allows the user to individually format the side wall of a 3D chart. Read-only.
- `BackWall As Walls  (read-only)`  
  Returns an object that allows the user to individually format the back wall of a 3D chart. Read-only Walls.
- `ChartStyle As Variant  (read/write)`  
  Returns or sets the chart style for the chart. Read/write Variant.
- `ShowDataLabelsOverMaximum As Boolean  (read/write)`  
  Returns or sets a value that indicates whether to show the data labels when the value is greater than the maximum value on the value axis. Read/write Boolean.
- `ChartData As ChartData  (read-only)`  
  Returns information about the linked or embedded data associated with a chart. Read-only ChartData.
- `Shapes As Shapes  (read-only)`  
  Returns a collection that represents all the shapes on the chart sheet. Read-only Shapes.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Application As Application  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft PowerPoint application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `AutoScaling As Boolean  (read/write)`  
  True if Microsoft Word scales a 3D chart so that it is closer in size to the equivalent 2D chart. The RightAngleAxes property must be True. Read/write Boolean.
- `ChartArea As ChartArea  (read-only)`  
  Returns the complete chart area for the chart. Read-only ChartArea.
- `ChartTitle As ChartTitle  (read-only)`  
  Returns the title of the specified chart. Read-only ChartTitle.
- `DepthPercent As Long  (read/write)`  
  Returns or sets the depth of a 3D chart as a percentage of the chart width (between 20 and 2000 percent). Read/write Long.
- `DisplayBlanksAs As XlDisplayBlanksAs  (read/write)`  
  Returns or sets the way that blank cells are plotted on a chart. Can be one of the XlDisplayBlanksAs constants. Read/write Long.
- `Elevation As Long  (read/write)`  
  Returns or sets the elevation, in degrees, of the 3D chart view. Read/write Long.
- `Floor As Floor  (read-only)`  
  Returns the floor of the 3D chart. Read-only Floor.
- `GapDepth As Long  (read/write)`  
  Returns or sets the distance, as a percentage of the marker width, between the data series in a 3D chart. Read/write Long.
- `HasAxis As Variant  (read/write)`  
  Returns or sets which axes exist on the chart. Read/write Variant.
- `HasLegend As Boolean  (read/write)`  
  True if the chart has a legend. Read/write Boolean.
- `HasTitle As Boolean  (read/write)`  
  True if the axis or chart has a visible title. Read/write Boolean.
- `HeightPercent As Long  (read/write)`  
  Returns or sets the height of a 3D chart as a percentage of the chart width (from 5 through 500 percent). Read/write Long.
- `Legend As Legend  (read-only)`  
  Returns the legend for the chart. Read-only Legend.
- `Name As String  (read/write)`  
  Read/write
- `Perspective As Long  (read/write)`  
  Returns or sets the perspective for the 3D chart view. Read/write Long.
- `PlotArea As PlotArea  (read-only)`  
  Returns the plot area of a chart. Read-only PlotArea.
- `PlotVisibleOnly As Boolean  (read/write)`  
  True if only visible cells are plotted. False if both visible and hidden cells are plotted. Read/write Boolean.
- `RightAngleAxes As Variant  (read/write)`  
  True if the chart axes are at right angles, independent of chart rotation or elevation. Read/write Boolean.
- `Rotation As Variant  (read/write)`  
  Returns or sets the rotation, in degrees, of the 3D chart view (the rotation of the plot area around the z-axis). Read/write Variant.
- `Walls As Walls  (read-only)`  
  Returns the walls of the 3D chart. Read-only Walls.
- `Format As ChartFormat  (read-only)`  
  Returns the ChartFormat object. Read-only.
- `ShowReportFilterFieldButtons As Boolean  (read/write)`  
  Returns or sets a value that indicates whether to display the report filter field buttons on a PivotChart. Read/write.
- `ShowLegendFieldButtons As Boolean  (read/write)`  
  Returns or sets a value that indicates whether to display legend field buttons on a PivotChart. Read/write.
- `ShowAxisFieldButtons As Boolean  (read/write)`  
  Returns or sets a value that indicates whether to display axis field buttons on a PivotChart. Read/write
- `ShowValueFieldButtons As Boolean  (read/write)`  
  Returns or sets a value that indicates whether to display the value field buttons on a PivotChart. Read/write.
- `ShowAllFieldButtons As Boolean  (read/write)`  
  Returns or sets a value that indicates whether to display all field buttons on a PivotChart. Read/write.
- `AlternativeText As String  (read/write)`  
  Returns or sets the alternative text associated with a shape in a Web presentation. Read/write.
- `Title As String  (read/write)`  
  Gets or sets a String that represents the title of the chart. Read/write.
- `CategoryLabelLevel As XlCategoryLabelLevel  (read/write)`  
  Returns or sets an XlCategoryLabel constant that specifies the source level of the chart category labels. Read/write.
- `SeriesNameLevel As XlSeriesNameLevel  (read/write)`  
  Returns or sets an XlSeriesNameLevel constant that specifies the source level of the series names. Read/write.
- `ChartColor As Variant  (read/write)`  
  Returns or sets an integer that represents the color scheme for the chart. Read/write.

## Methods (25)

- `ApplyDataLabels([Type As XlDataLabelsType], [LegendKey As Variant], [AutoText As Variant], [HasLeaderLines As Variant], [ShowSeriesName As Variant], [ShowCategoryName As Variant], [ShowValue As Variant], [ShowPercentage As Variant], [ShowBubbleSize As Variant], [Separator As Variant])`  
  Applies data labels to all the series in a chart.
    - `Type As XlDataLabelsType` (optional): One of the enumeration values that specifies the type of data label to apply. Can be one of the xlDataLabelsType constants.
    - `LegendKey As Variant` (optional): True to show the legend key next to the point. The default is False.
    - `AutoText As Variant` (optional): True if the object automatically generates appropriate text based on content.
    - `HasLeaderLines As Variant` (optional): For the Chart and Series objects, True if the series has leader lines.
    - `ShowSeriesName As Variant` (optional): True to enable the series name for the data label; otherwise, False.
    - `ShowCategoryName As Variant` (optional): True to enable the category name for the data label; otherwise, False.
    - `ShowValue As Variant` (optional): True to enable the value for the data label; otherwise, False.
    - `ShowPercentage As Variant` (optional): True to enable the percentage for the data label; otherwise, False.
    - `ShowBubbleSize As Variant` (optional): True to enable the bubble size for the data label; otherwise, False.
    - `Separator As Variant` (optional): The separator for the data label.
- `GetChartElement(X As Long, Y As Long, ElementID As Long, Arg1 As Long, Arg2 As Long)`  
  Returns information about the chart element at the specified x-coordinate and y-coordinate.
    - `X As Long` (required): The x-coordinate of the chart element.
    - `Y As Long` (required): The y-coordinate of the chart element.
    - `ElementID As Long` (required): When the method returns, this argument contains the XlChartItem value of the chart element at the specified coordinates. For more information, see the Remarks section.
    - `Arg1 As Long` (required): When the method returns, this argument contains information related to the chart element. For more information, see the Remarks section.
    - `Arg2 As Long` (required): When the method returns, this argument contains information related to the chart element. For more information, see the Remarks section.
- `SetSourceData(Source As String, [PlotBy As Variant])`  
  Sets the source data range for the chart.
    - `Source As String` (required): The address of the chart data range that contains the source data.
    - `PlotBy As Variant` (optional): Specifies the way the data will be plotted. Can be either of the following XlRowCol constants: xlColumns or xlRows.
- `SetBackgroundPicture(FileName As String)`  
  Sets the background graphic for a chart.
    - `FileName As String` (required): The name of the file for the graphic.
- `Paste([Type As Variant])`  
  Pastes chart data from the Clipboard into the chart.
    - `Type As Variant` (optional): Specifies the chart information to paste if a chart is on the Clipboard. Can be one of the following values:
- `SetDefaultChart(Name As Variant)`  
  Specifies the name of the chart template that Microsoft Word uses when it creates new charts.
    - `Name As Variant` (required): Specifies the name of the default chart template that Word uses when it creates new charts. This name can be set to either the name of a user-defined chart template in the gallery or a special XlChartGallery constant, xlBuiltIn, to specify a built-in chart template.
- `ApplyChartTemplate(FileName As String)`  
  Applies a standard or custom chart type to a chart.
    - `FileName As String` (required): The file name for a chart template.
- `SaveChartTemplate(FileName As String)`  
  Saves a custom chart template to the list of available chart templates.
    - `FileName As String` (required): The name of the chart template.
- `ClearToMatchStyle()`  
  Clears the chart elements formatting to automatic.
- `ApplyLayout(Layout As Long, [ChartType As Variant])`  
  Applies the layouts shown in the Ribbon.
    - `Layout As Long` (required): The type of layout. The type of layout is denoted by a number from 1 to 10.
    - `ChartType As Variant` (optional): An XlChartType constant that represents the type of chart.
- `Refresh()`  
  Causes the specified chart to be redrawn immediately.
- `Axes([Type As Variant], [AxisGroup As XlAxisGroup]) As Object`  
  Returns a collection of axes on the chart.
    - `Type As Variant` (optional): The axis to return. Can be one of the following XlAxisType constants: xlValue, xlCategory, or xlSeriesAxis (xlSeriesAxis is valid only for 3D charts).
    - `AxisGroup As XlAxisGroup` (optional): One of the enumeration values that specifies the axis group. The default is xlPrimary.
- `ChartGroups([Index As Variant]) As Object`  
  Returns an object that represents either a single chart group or a collection of all the chart groups in the chart.
    - `Index As Variant` (optional): The chart group number. If specified, a single ChartGroup object is returned. If omitted, a ChartGroups object is returned which contains a collection of every ChartGroup object for that chart.
- `ChartWizard([Source As Variant], [Gallery As Variant], [Format As Variant], [PlotBy As Variant], [CategoryLabels As Variant], [SeriesLabels As Variant], [HasLegend As Variant], [Title As Variant], [CategoryTitle As Variant], [ValueTitle As Variant], [ExtraTitle As Variant])`  
  Modifies the properties of the given chart. Use this method to quickly format a chart without setting all the individual properties. This method is noninteractive, and it changes only the specified properties.
    - `Source As Variant` (optional): The range that contains the source data for the new chart. If this argument is omitted, Word edits the active chart sheet or the selected chart on the active worksheet.
    - `Gallery As Variant` (optional): One of the XlChartType constants that specifies the chart type.
    - `Format As Variant` (optional): The option number for the built-in autoformats. Can be a number from 1 through 10, depending on the gallery type. If this argument is omitted, Word chooses a default value based on the gallery type and data source.
    - `PlotBy As Variant` (optional): Specifies whether the data for each series is in rows or columns. Can be one of the following XlRowCol constants: xlRows or xlColumns.
    - `CategoryLabels As Variant` (optional): An integer that specifies the number of rows or columns within the source range that contain category labels. Allowed values are from 0 (zero) through one less than the maximum number of the corresponding categories or series.
    - `SeriesLabels As Variant` (optional): An integer that specifies the number of rows or columns within the source range that contain series labels. Allowed values are from 0 (zero) through one less than the maximum number of the corresponding categories or series.
    - `HasLegend As Variant` (optional): True to include a legend.
    - `Title As Variant` (optional): The chart title text.
    - `CategoryTitle As Variant` (optional): The category axis title text.
    - `ValueTitle As Variant` (optional): The value axis title text.
    - `ExtraTitle As Variant` (optional): The series axis title for 3D charts or the second value axis title for 2D charts.
- `Copy([Before As Variant], [After As Variant])`  
  Not supported for this object.
    - `Before As Variant` (optional): Not supported for this object.
    - `After As Variant` (optional): Not supported for this object.
- `CopyPicture([Appearance As XlPictureAppearance], [Format As XlCopyPictureFormat], [Size As XlPictureAppearance])`  
  Copies the selected object to the Clipboard as a picture.
    - `Appearance As XlPictureAppearance` (optional): One of the enumeration values that specifies how the picture should be copied. The default is xlScreen.
    - `Format As XlCopyPictureFormat` (optional): One of the enumeration values that specifies the format of the picture. The default is xlPicture.
    - `Size As XlPictureAppearance` (optional): One of the enumeration values that specifies the size of the copied picture when the object is a chart on a chart sheet (not embedded on a worksheet). The default is xlPrinter.
- `Delete()`  
  Deletes the object.
- `Export(FileName As String, [FilterName As Variant], [Interactive As Variant]) As Boolean`  
  Exports the chart in a graphic format.
    - `FileName As String` (required): The name of the exported file.
    - `FilterName As Variant` (optional): The language-independent name of the graphic filter as it appears in the registry, for example, PNG or GIF. PNG is the default if no value is supplied.
    - `Interactive As Variant` (optional): N/A
- `Select([Replace As Variant])`  
  Selects the object.
    - `Replace As Variant` (optional): True to replace the current selection with the specified object. False to extend the current selection to include any previously selected objects and the specified object. This parameter is used only with sheets.
- `SeriesCollection([Index As Variant]) As Object`  
  Returns all the series in the chart.
- `SetElement(Element As MsoChartElementType)`  
  Sets chart elements on a chart. Read/write MsoChartElementType.
    - `Element As MsoChartElementType` (required): One of the enumeration values that specifies the chart element type.
- `FullSeriesCollection([Index As Variant]) As Object`  
  Returns the collection of all the series in the specified chart, or the specified series.
- `ClearToMatchColorStyle()`  
  Clears all colors on the specified chart that don't follow the color style applied to the chart.
- `GetProperty(Id As String) As Variant`
- `SetProperty(Id As String, Value As Variant)`
