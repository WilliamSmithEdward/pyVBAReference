# IMsoChart

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C1709-0000-0000-C000-000000000046}  

## Properties (46)

- `Parent As Object  (read-only)`
- `HasTitle As Boolean  (read/write)`
- `ChartTitle As IMsoChartTitle  (read-only)`
- `DepthPercent As Long  (read/write)`
- `Elevation As Long  (read/write)`
- `GapDepth As Long  (read/write)`
- `HeightPercent As Long  (read/write)`
- `Perspective As Long  (read/write)`
- `RightAngleAxes As Variant  (read/write)`
- `Rotation As Variant  (read/write)`
- `DisplayBlanksAs As XlDisplayBlanksAs  (read/write)`
- `ChartGroups As Object  (read-only)`
- `ChartType As XlChartType  (read/write)`
- `HasDataTable As Boolean  (read/write)`
- `PlotBy As XlRowCol  (read/write)`
- `HasLegend As Boolean  (read/write)`
- `Legend As IMsoLegend  (read-only)`
- `HasAxis As Variant  (read/write)`
- `Walls As IMsoWalls  (read-only)`
- `Floor As IMsoFloor  (read-only)`
- `PlotArea As IMsoPlotArea  (read-only)`
- `PlotVisibleOnly As Boolean  (read/write)`
- `ChartArea As IMsoChartArea  (read-only)`
- `AutoScaling As Boolean  (read/write)`
- `DataTable As IMsoDataTable  (read-only)`
- `BarShape As XlBarShape  (read/write)`
- `SideWall As IMsoWalls  (read-only)`
- `BackWall As IMsoWalls  (read-only)`
- `ChartStyle As Variant  (read/write)`
- `PivotLayout As Object  (read-only)`
- `ShowDataLabelsOverMaximum As Boolean  (read/write)`
- `ChartData As IMsoChartData  (read-only)`
- `Format As IMsoChartFormat  (read-only)`
- `Shapes As Shapes  (read-only)`
- `Application As Object  (read-only)`
- `Creator As Long  (read-only)`
- `ShowReportFilterFieldButtons As Boolean  (read/write)`
- `ShowLegendFieldButtons As Boolean  (read/write)`
- `ShowAxisFieldButtons As Boolean  (read/write)`
- `ShowValueFieldButtons As Boolean  (read/write)`
- `ShowAllFieldButtons As Boolean  (read/write)`
- `CategoryLabelLevel As XlCategoryLabelLevel  (read/write)`
- `SeriesNameLevel As XlSeriesNameLevel  (read/write)`
- `ChartColor As Variant  (read/write)`
- `ShowExpandCollapseEntireFieldButtons As Boolean  (read/write)`
- `DisplayValueNotAvailableAsBlank As Boolean  (read/write)`

## Methods (25)

- `SeriesCollection([Index As Variant]) As Object`
- `ApplyDataLabels([Type As XlDataLabelsType], [IMsoLegendKey As Variant], [AutoText As Variant], [HasLeaderLines As Variant], [ShowSeriesName As Variant], [ShowCategoryName As Variant], [ShowValue As Variant], [ShowPercentage As Variant], [ShowBubbleSize As Variant], [Separator As Variant])`
- `ApplyCustomType(ChartType As XlChartType, [TypeName As Variant])`
- `GetChartElement(x As Long, y As Long, ElementID As Long, Arg1 As Long, Arg2 As Long)`
- `SetSourceData(Source As String, [PlotBy As Variant])`
- `Axes([Type As Variant], [AxisGroup As XlAxisGroup]) As Object`
- `AutoFormat(rGallery As Long, [varFormat As Variant])`
- `ChartWizard([varSource As Variant], [varGallery As Variant], [varFormat As Variant], [varPlotBy As Variant], [varCategoryLabels As Variant], [varSeriesLabels As Variant], [varHasLegend As Variant], [varTitle As Variant], [varCategoryTitle As Variant], [varValueTitle As Variant], [varExtraTitle As Variant])`
- `CopyPicture([Appearance As Long], [Format As Long], [Size As Long])`
- `Export(bstr As String, [varFilterName As Variant], [varInteractive As Variant]) As Boolean`
- `SetDefaultChart(varName As Variant)`
- `ApplyChartTemplate(bstrFileName As String)`
- `SaveChartTemplate(bstrFileName As String)`
- `ClearToMatchStyle()`
- `RefreshPivotTable()`
- `ApplyLayout(Layout As Long, [varChartType As Variant])`
- `Refresh()`
- `SetElement(RHS As MsoChartElementType)`
- `Delete() As Variant`
- `Copy() As Variant`
- `Select([Replace As Variant]) As Variant`
- `FullSeriesCollection([Index As Variant]) As Object`
- `ClearToMatchColorStyle()`
- `SetProperty(bstrId As String, Value As Variant)`
- `GetProperty(bstrId As String) As Variant`
