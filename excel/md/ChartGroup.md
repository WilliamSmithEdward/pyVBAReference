# ChartGroup

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020859-0000-0000-C000-000000000046}  

Represents one or more series plotted in a chart with the same format.

**Remarks:** A chart contains one or more chart groups, each chart group contains one or more Series objects, and each series contains one or more Points objects. For example, a single chart might contain both a line chart group that contains all the series plotted with the line chart format, and a bar chart group that contains all the series plotted with the bar chart format. The ChartGroup object is a member of the ChartGroups collection. Use ChartGroups (_index_), where _index_ is the chart-group index number, to return a single ChartGroup object. Because the index number for a particular chart group can change if the chart format used for that group is changed, it may be easier to use one of the named chart group shortcut methods to return a particular chart group. The PieGroups method returns the collection of pie chart groups in a chart, the LineGroups method returns the collection of line chart groups, and so on. Each of these methods can be used with an index number to return a single ChartGroup object, or without an index number to return a ChartGroups collection.

**Example:**

```vba
Charts(1).ChartGroups(1).HasDropLines = True
```

## Properties (35)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `AxisGroup As XlAxisGroup  (read/write)`  
  Returns or sets the group for the specified chart. Read/write.
- `DoughnutHoleSize As Long  (read/write)`  
  Returns or sets the size of the hole in a doughnut chart group. The hole size is expressed as a percentage of the chart size, between 10 and 90 percent. Read/write Long.
- `DownBars As DownBars  (read-only)`  
  Returns a DownBars object that represents the down bars on a line chart. Applies only to line charts. Read-only.
- `DropLines As DropLines  (read-only)`  
  Returns a DropLines object that represents the drop lines for a series on a line chart or area chart. Applies only to line charts or area charts. Read-only.
- `FirstSliceAngle As Long  (read/write)`  
  Returns or sets the angle of the first pie-chart or doughnut-chart slice, in degrees (clockwise from vertical). Applies only to pie, 3D pie, and doughnut charts. Can be a value from 0 through 360. Read/write Long.
- `GapWidth As Long  (read/write)`  
  Bar and Column charts: Returns or sets the space between bar or column clusters, as a percentage of the bar or column width.
- `HasDropLines As Boolean  (read/write)`  
  True if the line chart or area chart has drop lines. Applies only to line and area charts. Read/write Boolean.
- `HasHiLoLines As Boolean  (read/write)`  
  True if the line chart has high-low lines. Applies only to line charts. Read/write Boolean.
- `HasRadarAxisLabels As Boolean  (read/write)`  
  True if a radar chart has axis labels. Applies only to radar charts. Read/write Boolean.
- `HasSeriesLines As Boolean  (read/write)`  
  True if a stacked column chart or bar chart has series lines, or if a Pie of Pie chart or Bar of Pie chart has connector lines between the two sections. Applies only to 2D stacked bar, 2D stacked column, Pie of Pie, or Bar of Pie charts. Read/write Boolean.
- `HasUpDownBars As Boolean  (read/write)`  
  True if a line chart has up and down bars. Applies only to line charts. Read/write Boolean.
- `HiLoLines As HiLoLines  (read-only)`  
  Returns a HiLoLines object that represents the high-low lines for a series on a line chart. Applies only to line charts. Read-only.
- `Index As Long  (read-only)`  
  Returns a Long value that represents the index number of the object within the collection of similar objects.
- `Overlap As Long  (read/write)`  
  Specifies how bars and columns are positioned. Can be a value between -100 and 100. Applies only to 2D bar and 2D column charts. Read/write Long.
- `RadarAxisLabels As TickLabels  (read-only)`  
  Returns a TickLabels object that represents the radar axis labels for the specified chart group. Read-only.
- `SeriesLines As SeriesLines  (read-only)`  
  Returns a SeriesLines object that represents the series lines for a 2D stacked bar, 2D stacked column, Pie of Pie, or Bar of Pie chart. Read-only.
- `UpBars As UpBars  (read-only)`  
  Returns an UpBars object that represents the up bars on a line chart. Applies only to line charts. Read-only.
- `VaryByCategories As Boolean  (read/write)`  
  True if Microsoft Excel assigns a different color or pattern to each data marker. The chart must contain only one series. Read/write Boolean.
- `SizeRepresents As XlSizeRepresents  (read/write)`  
  Returns or sets what the bubble size represents on a bubble chart. Can be either of the following XlSizeRepresents constants: xlSizeIsArea or xlSizeIsWidth. Read/write Long.
- `BubbleScale As Long  (read/write)`  
  Returns or sets the scale factor for bubbles in the specified chart group. Can be an integer value from 0 (zero) to 300, corresponding to a percentage of the default size. Applies only to bubble charts. Read/write Long.
- `ShowNegativeBubbles As Boolean  (read/write)`  
  True if negative bubbles are shown for the chart group. Valid only for bubble charts. Read/write Boolean.
- `SplitType As XlChartSplitType  (read/write)`  
  Returns or sets the way the two sections of either a Pie of Pie chart or a Bar of Pie chart are split. Read/write XlChartSplitType.
- `SplitValue As Variant  (read/write)`  
  Returns or sets the threshold value separating the two sections of either a Pie of Pie chart or a Bar of Pie chart. Read/write Variant.
- `SecondPlotSize As Long  (read/write)`  
  Returns or sets the size of the secondary section of either a Pie of Pie chart or a Bar of Pie chart, as a percentage of the size of the primary pie. Can be a value from 5 to 200. Read/write Long.
- `Has3DShading As Boolean  (read/write)`  
  Returns or sets the 3D shading property of a ChartGroup object. Read/write Boolean.
- `BinsType As XlBinsType  (read/write)`  
  Specifies how the horizontal axis of the histogram chart is formatted, by bins type. Read/write XlBinsType.
- `BinWidthValue As Double  (read/write)`  
  Specifies the number of points in each range. Read/write Double.
- `BinsCountValue As Long  (read/write)`  
  Specifies the number of bins in the histogram chart. Read/write Long.
- `BinsOverflowEnabled As Boolean  (read/write)`  
  Specifies whether a bin for values above the BinsOverflowValue property is enabled. Read/write Boolean.
- `BinsOverflowValue As Double  (read/write)`  
  If a BinsOverflowEnabled property is True, specifies the value above which an overflow bin is displayed. Read/write Double.
- `BinsUnderflowEnabled As Boolean  (read/write)`  
  Specifies whether a bin for values below the BinsUnderflowValue property is enabled. Read/write Boolean.
- `BinsUnderflowValue As Double  (read/write)`  
  If a BinsUnderflowEnabled property is True, specifies the value below which an underflow bin is displayed. Read/write Double.

## Methods (3)

- `SeriesCollection([Index As Variant]) As Object`  
  Returns an object that represents either a single series (a Series object) or a collection of all the series (a SeriesCollection collection) in the chart or chart group.
    - `Index As Variant` (optional): The name or number of the series.
- `FullCategoryCollection([Index As Variant]) As Object`  
  Returns an object that represents a collection of all the visible and filtered categories (a CategoryCollection collection) in the chart group.
    - `Index As Variant` (optional): The name or number of the categories.
- `CategoryCollection([Index As Variant]) As Object`  
  Returns an object that represents a collection of all the visible categories (a CategoryCollection collection) in the chart group.
    - `Index As Variant` (optional): The name or number of the categories.
