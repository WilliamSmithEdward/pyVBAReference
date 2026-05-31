# Series

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002086B-0000-0000-C000-000000000046}  

## Properties (52)

- `Application As Application  (read-only)`
- `Creator As XlCreator  (read-only)`
- `Parent As Object  (read-only)`
- `AxisGroup As XlAxisGroup  (read/write)`
- `ErrorBars As ErrorBars  (read-only)`
- `Explosion As Long  (read/write)`
- `Formula As String  (read/write)`
- `FormulaLocal As String  (read/write)`
- `FormulaR1C1 As String  (read/write)`
- `FormulaR1C1Local As String  (read/write)`
- `HasDataLabels As Boolean  (read/write)`
- `HasErrorBars As Boolean  (read/write)`
- `InvertIfNegative As Boolean  (read/write)`
- `MarkerBackgroundColor As Long  (read/write)`
- `MarkerBackgroundColorIndex As XlColorIndex  (read/write)`
- `MarkerForegroundColor As Long  (read/write)`
- `MarkerForegroundColorIndex As XlColorIndex  (read/write)`
- `MarkerSize As Long  (read/write)`
- `MarkerStyle As XlMarkerStyle  (read/write)`
- `Name As String  (read/write)`
- `PictureType As XlChartPictureType  (read/write)`
- `PlotOrder As Long  (read/write)`
- `Smooth As Boolean  (read/write)`
- `Type As Long  (read/write)`
- `ChartType As XlChartType  (read/write)`
- `Values As Variant  (read/write)`
- `XValues As Variant  (read/write)`
- `BubbleSizes As Variant  (read/write)`
- `BarShape As XlBarShape  (read/write)`
- `ApplyPictToSides As Boolean  (read/write)`
- `ApplyPictToFront As Boolean  (read/write)`
- `ApplyPictToEnd As Boolean  (read/write)`
- `Has3DEffect As Boolean  (read/write)`
- `Shadow As Boolean  (read/write)`
- `HasLeaderLines As Boolean  (read/write)`
- `LeaderLines As LeaderLines  (read-only)`
- `PictureUnit2 As Double  (read/write)`
- `Format As ChartFormat  (read-only)`
- `PlotColorIndex As Long  (read-only)`
- `InvertColor As Long  (read/write)`
- `InvertColorIndex As Long  (read/write)`
- `IsFiltered As Boolean  (read/write)`
- `ParentDataLabelOption As XlParentDataLabelOptions  (read/write)`
- `QuartileCalculationInclusiveMedian As Boolean  (read/write)`
- `ValueSortOrder As XlValueSortOrder  (read/write)`
- `GeoProjectionType As XlGeoProjectionType  (read/write)`
- `GeoMappingLevel As XlGeoMappingLevel  (read/write)`
- `RegionLabelOption As XlRegionLabelOptions  (read/write)`
- `SeriesColorGradientStyle As XlSeriesColorGradientStyle  (read/write)`
- `SeriesColorMinGradientStop As ChartSeriesGradientStopData  (read-only)`
- `SeriesColorMidGradientStop As ChartSeriesGradientStopData  (read-only)`
- `SeriesColorMaxGradientStop As ChartSeriesGradientStopData  (read-only)`

## Methods (12)

- `ClearFormats() As Variant`
- `Copy() As Variant`
- `DataLabels([Index As Variant]) As Object`
- `Delete() As Variant`
- `ErrorBar(Direction As XlErrorBarDirection, Include As XlErrorBarInclude, Type As XlErrorBarType, [Amount As Variant], [MinusValues As Variant]) As Variant`
- `Paste() As Variant`
- `Points([Index As Variant]) As Object`
- `Select() As Variant`
- `Trendlines([Index As Variant]) As Object`
- `ApplyDataLabels([Type As XlDataLabelsType], [LegendKey As Variant], [AutoText As Variant], [HasLeaderLines As Variant], [ShowSeriesName As Variant], [ShowCategoryName As Variant], [ShowValue As Variant], [ShowPercentage As Variant], [ShowBubbleSize As Variant], [Separator As Variant]) As Variant`
- `SetProperty(ID As String, Value As Variant)`
- `GetProperty(ID As String) As Variant`
