# ChartGroup

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {86488FB4-9633-4C93-8057-FC1FA7A847AE}  

Represents one or more series plotted in a chart with the same format.

**Remarks:** A chart contains one or more chart groups, each chart group contains one or more Series objects, and each series contains one or more Points objects. For example, a single chart might contain both a line chart group, which contains all the series plotted with the line chart format, and a bar chart group, which contains all the series plotted with the bar chart format. The ChartGroup object is a member of the ChartGroups collection. Use ChartGroups (_index_), where _index_ is the chart group index number, to return a single ChartGroup object.

**Example:**

```vba
With ActiveDocument.InlineShapes(1).Chart
 .ChartGroups(1).HasDropLines = True
End With
```

## Properties (35)

- `AxisGroup As XlAxisGroup  (read/write)`  
  Returns the type of axis group. Read/write XlAxisGroup.
- `DoughnutHoleSize As Long  (read/write)`  
  Returns or sets the size of the hole in a doughnut chart group. Read/write Long.
- `DownBars As DownBars  (read-only)`  
  Returns the down bars on a line chart. Read-only DownBars.
- `DropLines As DropLines  (read-only)`  
  Returns the drop lines for a series on a line chart or area chart. Read-only DropLines.
- `FirstSliceAngle As Long  (read/write)`  
  Returns or sets the angle, in degrees (clockwise from vertical), of the first pie-chart or doughnut-chart slice. Read/write Long.
- `GapWidth As Long  (read/write)`  
  For bar and column charts, returns or sets the space, as a percentage of the bar or column width, between bar or column clusters. For pie-of-pie and bar-of-pie charts, returns or sets the space between the primary and secondary sections of the chart. Read/write Long.
- `HasDropLines As Boolean  (read/write)`  
  True if the line chart or area chart has drop lines. Read/write Boolean.
- `HasHiLoLines As Boolean  (read/write)`  
  True if the line chart has high-low lines. Read/write Boolean.
- `HasRadarAxisLabels As Boolean  (read/write)`  
  True if a radar chart has axis labels. Read/write Boolean.
- `HasSeriesLines As Boolean  (read/write)`  
  True if a stacked column chart or bar chart has series lines or if a pie-of-pie chart or bar-of-pie chart has connector lines between the two sections. Read/write Boolean.
- `HasUpDownBars As Boolean  (read/write)`  
  True if a line chart has up and down bars. Read/write Boolean.
- `HiLoLines As HiLoLines  (read-only)`  
  Returns the high-low lines for a series on a line chart. Read-only HiLoLines.
- `Index As Long  (read-only)`  
  Returns the index number of the object within the collection of similar objects. Read-only Long.
- `Overlap As Long  (read/write)`  
  Specifies how bars and columns are positioned. Read/write Long.
- `RadarAxisLabels As TickLabels  (read-only)`  
  Returns the radar axis labels for the specified chart group. Read-only TickLabels.
- `SeriesLines As SeriesLines  (read-only)`  
  Returns the series lines for a 2D stacked bar, 2D stacked column, pie-of-pie, or bar-of-pie chart. Read-only SeriesLines.
- `UpBars As UpBars  (read-only)`  
  Returns the up bars on a line chart. Read-only UpBars.
- `VaryByCategories As Boolean  (read/write)`  
  True if Microsoft Word assigns a different color or pattern to each data marker. Read/write Boolean.
- `SizeRepresents As XlSizeRepresents  (read/write)`  
  Returns or sets what the bubble size represents on a bubble chart. Read/write Long.
- `BubbleScale As Long  (read/write)`  
  Returns or sets the scale factor for bubbles in the specified chart group. Read/write Long.
- `ShowNegativeBubbles As Boolean  (read/write)`  
  True if negative bubbles are shown for the chart group. Read/write Boolean.
- `SplitType As XlChartSplitType  (read/write)`  
  Returns or sets the way the two sections of either a pie-of-pie chart or a bar-of-pie chart are split. Read/write XlChartSplitType.
- `SplitValue As Variant  (read/write)`  
  Returns or sets the threshold value separating the two sections of either a pie-of-pie chart or a bar-of-pie chart. Read/write Variant.
- `SecondPlotSize As Long  (read/write)`  
  Returns or sets the size, as a percentage of the primary pie, of the secondary section of either a pie-of-pie chart or a bar-of-pie chart. Read/write Long.
- `Has3DShading As Boolean  (read/write)`  
  True if a chart group has three-dimensional shading. Read/write Boolean.
- `Application As Object  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft Word application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `BinsType As XlBinsType  (read/write)`  
  Specifies how the horizontal axis of the histogram chart is formatted, by bins type. Read/write XlBinsType.
- `BinWidthValue As Double  (read/write)`  
  Specifies the number of points in each range. Read/write Double.
- `BinsCountValue As Long  (read/write)`  
  Specifies the number of bins in the histogram chart. Read/write Long.
- `BinsOverflowEnabled As Boolean  (read/write)`  
  Specifies whether a bin for values above the BinsOverflowValue is enabled. Read/write Boolean.
- `BinsOverflowValue As Double  (read/write)`  
  If an BinsOverflowEnabled is True, specifies the value above which an overflow bin is displayed. Read/write Double.
- `BinsUnderflowEnabled As Boolean  (read/write)`  
  Specifies whether a bin for values below the BinsUnderflowValue is enabled. Read/write Boolean.
- `BinsUnderflowValue As Double  (read/write)`  
  If an BinsUnderflowEnabled is True, specifies the value below which an underflow bin is displayed. Read/write Double.

## Methods (3)

- `SeriesCollection([Index As Variant]) As Object`  
  Returns all the series in the chart group.
- `CategoryCollection([Index As Variant]) As Object`  
  Returns all the visible categories in the chart group, or the specified visible category.
    - `Index As Variant` (optional): The index number in the visible category collection of the category to return.
- `FullCategoryCollection([Index As Variant]) As Object`  
  Returns all the categories in the chart group, or the specified category, whether visible or filtered out.
    - `Index As Variant` (optional): The index number in the full category collection of the category to return.
