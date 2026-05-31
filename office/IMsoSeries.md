# IMsoSeries

**Type:** Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C170B-0000-0000-C000-000000000046}  

## Properties (56)

- `Parent As HRESULT  (read-only)`
- `AxisGroup As HRESULT  (read/write)`
- `Border As HRESULT  (read-only)`
- `ErrorBars As HRESULT  (read-only)`
- `Explosion As HRESULT  (read/write)`
- `Formula As HRESULT  (read/write)`
- `FormulaLocal As HRESULT  (read/write)`
- `FormulaR1C1 As HRESULT  (read/write)`
- `FormulaR1C1Local As HRESULT  (read/write)`
- `HasDataLabels As HRESULT  (read/write)`
- `HasErrorBars As HRESULT  (read/write)`
- `Interior As HRESULT  (read-only)`
- `Fill As HRESULT  (read-only)`
- `InvertIfNegative As HRESULT  (read/write)`
- `MarkerBackgroundColor As HRESULT  (read/write)`
- `MarkerBackgroundColorIndex As HRESULT  (read/write)`
- `MarkerForegroundColor As HRESULT  (read/write)`
- `MarkerForegroundColorIndex As HRESULT  (read/write)`
- `MarkerSize As HRESULT  (read/write)`
- `MarkerStyle As HRESULT  (read/write)`
- `Name As HRESULT  (read/write)`
- `PictureType As HRESULT  (read/write)`
- `PictureUnit As HRESULT  (read/write)`
- `PlotOrder As HRESULT  (read/write)`
- `Smooth As HRESULT  (read/write)`
- `Type As HRESULT  (read/write)`
- `ChartType As HRESULT  (read/write)`
- `Values As HRESULT  (read/write)`
- `XValues As HRESULT  (read/write)`
- `BubbleSizes As HRESULT  (read/write)`
- `BarShape As HRESULT  (read/write)`
- `ApplyPictToSides As HRESULT  (read/write)`
- `ApplyPictToFront As HRESULT  (read/write)`
- `ApplyPictToEnd As HRESULT  (read/write)`
- `Has3DEffect As HRESULT  (read/write)`
- `Shadow As HRESULT  (read/write)`
- `HasLeaderLines As HRESULT  (read/write)`
- `LeaderLines As HRESULT  (read-only)`
- `Format As HRESULT  (read-only)`
- `Application As HRESULT  (read-only)`
- `Creator As HRESULT  (read-only)`
- `PictureUnit2 As HRESULT  (read/write)`
- `PlotColorIndex As HRESULT  (read-only)`
- `InvertColor As HRESULT  (read/write)`
- `InvertColorIndex As HRESULT  (read/write)`
- `IsFiltered As HRESULT  (read/write)`
- `ParentDataLabelOption As HRESULT  (read/write)`
- `QuartileCalculationInclusiveMedian As HRESULT  (read/write)`
- `ValueSortOrder As HRESULT  (read/write)`
- `GeoProjectionType As HRESULT  (read/write)`
- `GeoMappingLevel As HRESULT  (read/write)`
- `RegionLabelOption As HRESULT  (read/write)`
- `SeriesColorGradientStyle As HRESULT  (read/write)`
- `SeriesColorMinGradientStop As HRESULT  (read-only)`
- `SeriesColorMidGradientStop As HRESULT  (read-only)`
- `SeriesColorMaxGradientStop As HRESULT  (read-only)`

## Methods (13)

- `ClearFormats(RHS As Variant)`
- `Copy(RHS As Variant)`
- `DataLabels([Index As Variant], RHS As Object)`
- `Delete(RHS As Variant)`
- `ErrorBar(Direction As XlErrorBarDirection, Include As XlErrorBarInclude, Type As XlErrorBarType, [Amount As Variant], [MinusValues As Variant], RHS As Variant)`
- `Paste(RHS As Variant)`
- `Points([Index As Variant], RHS As Object)`
- `Select(RHS As Variant)`
- `Trendlines([Index As Variant], RHS As Object)`
- `ApplyCustomType(ChartType As XlChartType)`
- `ApplyDataLabels([Type As XlDataLabelsType], [IMsoLegendKey As Variant], [AutoText As Variant], [HasLeaderLines As Variant], [ShowSeriesName As Variant], [ShowCategoryName As Variant], [ShowValue As Variant], [ShowPercentage As Variant], [ShowBubbleSize As Variant], [Separator As Variant], RHS As Variant)`
- `SetProperty(bstrId As String, Value As Variant)`
- `GetProperty(bstrId As String, pValue As Variant)`
