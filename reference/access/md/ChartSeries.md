# ChartSeries

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {D91560D6-F71B-41F8-8D32-285062068258}  

Represents a series of values in the specified chart.

## Properties (27)

- `Name As String  (read-only)`  
  Returns the name of a ChartSeries instance based on the value of its associated field. Read-only String.
- `PlotSeriesOn As AcValueAxis  (read/write)`  
  Returns or sets the axis to use for plotting the values of a series. Read/write AcValueAxis.
- `DisplayDataLabel As Boolean  (read/write)`  
  True if the data for a series is displayed in a label in addition to chart visualizations. Read/write Boolean.
- `FillColor As Long  (read/write)`  
  Returns or sets the fill color of a series visualization. Read/write String.
- `BorderColor As Long  (read/write)`  
  Returns or sets the border color of a series visualization. Read/write String.
- `TrendlineOptions As AcTrendlineOptions  (read/write)`  
  Returns or sets the type of trendline to render for a series. Read/write AcTrendlineOptions.
- `TrendlineName As String  (read/write)`  
  Returns or sets the name of the series trendline if the trendline is visible. Read/write String.
- `ComboChartType As AcChartType  (read/write)`  
  Returns or sets the chart type for the specified series. Read/write AcChartType.
- `DisplayName As String  (read/write)`  
  Returns or sets the legend display name of a series. Read/write String.
- `DashType As AcDashType  (read/write)`  
  Returns or sets the line dash style of a series when the effective ChartType is acChartLine, acChartLineStacked, or acChartLineStacked100. Read/write AcDashType.
- `LineWeight As Long  (read/write)`  
  Returns or sets the line weight (thickness) for a series when its effective chart type is line-based. Read/write Long.
- `MissingDataPolicy As AcMissingDataPolicy  (read/write)`  
  Returns or sets the plotting strategy of a series when its chart type is Line and values are missing. Read/write AcMissingDataPolicy.
- `MarkerType As AcMarkerType  (read/write)`  
  Returns or sets the marker shape to use for data points when the effective chart type of a series is line-based. Read/write AcMarkerType.
- `GridlinesType As AcGridlineType  (read/write)`  
  Returns or sets the gridlines type for a Modern Chart (only available in Current Channel). Read/write AcGridlineType.
- `DisplayBoxWhiskerDataPoints As Boolean  (read/write)`  
  True if data points are displayed for a Box and Whisker Modern Chart (only available in Current Channel). Read/write Boolean.
- `DisplayBoxWhiskerMeanMarker As Boolean  (read/write)`  
  True if the mean marker is displayed for a Box and Whisker Modern Chart (only available in Current Channel). Read/write Boolean.
- `DataLabelPosition As AcDataLabelPosition  (read/write)`  
  Returns or sets the data label position for Pie, Arc, and Doughnut Modern Charts (only available in Current Channel). Read/write AcDataLabelPosition.
- `WordCloudShape As AcWordCloudShape  (read/write)`  
  Returns or sets the shape for a Word Cloud Modern Chart (only available in Current Channel). Read/write AcWordCloudShape.
- `WordCloudWordOrientation As AcWordCloudWordOrientation  (read/write)`  
  Returns or sets the word orientation for a Word Cloud Modern Chart (only available in Current Channel). Read/write AcWordCloudWordOrientation.
- `ShowWaterfallConnectorLines As Boolean  (read/write)`  
  True if connector lines are displayed for a Waterfall Modern Chart (only available in Current Channel). Read/write Boolean.
- `ShowWaterfallTotal As Boolean  (read/write)`  
  True if the total is displayed for a Waterfall Modern Chart (only available in Current Channel). Read/write Boolean.
- `ShowFunnelPercentages As Boolean  (read/write)`  
  True if percentages are displayed for a Funnel Modern Chart (only available in Current Channel). Read/write Boolean.
- `PercentageDataLabelDecimalPlaces As AcPercentageDataLabelDecimalPlaces  (read/write)`  
  Returns or sets the demimal places for percentage data labels - only applicable to Pie, Arc, and Doughnut Modern Charts (only available in Current Channel). Read/write AcPercentageDataLabelDecimalPlaces.
- `DataLabelDisplayFormat As AcDataLabelDisplayFormat  (read/write)`  
  Returns or sets data label format for Pie, Arc, and Doughnut Modern Charts (only available in Current Channel). Read/write AcDataLabelDisplayFormat.
- `SortOrderType As AcSortOrderType  (read/write)`  
  Returns or sets the sort order type for series in a Modern Chart (only available in Current Channel). Read/write AcSortOrderType.
- `GridlinesColor As Long  (read/write)`  
  Returns or sets the gridlines color for Modern Charts (only available in Current Channel). Read/write String.
- `ParetoLineColor As Long  (read/write)`  
  Returns or sets the line color for a Pareto Modern Chart (only available in Current Channel). Read/write String.
