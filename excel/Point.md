# Point

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002086A-0000-0000-C000-000000000046}  

## Properties (28)

- `Application As Application  (read-only)`
- `Creator As XlCreator  (read-only)`
- `Parent As Object  (read-only)`
- `DataLabel As DataLabel  (read-only)`
- `Explosion As Long  (read/write)`
- `HasDataLabel As Boolean  (read/write)`
- `InvertIfNegative As Boolean  (read/write)`
- `MarkerBackgroundColor As Long  (read/write)`
- `MarkerBackgroundColorIndex As XlColorIndex  (read/write)`
- `MarkerForegroundColor As Long  (read/write)`
- `MarkerForegroundColorIndex As XlColorIndex  (read/write)`
- `MarkerSize As Long  (read/write)`
- `MarkerStyle As XlMarkerStyle  (read/write)`
- `PictureType As XlChartPictureType  (read/write)`
- `ApplyPictToSides As Boolean  (read/write)`
- `ApplyPictToFront As Boolean  (read/write)`
- `ApplyPictToEnd As Boolean  (read/write)`
- `Shadow As Boolean  (read/write)`
- `SecondaryPlot As Boolean  (read/write)`
- `Has3DEffect As Boolean  (read/write)`
- `PictureUnit2 As Double  (read/write)`
- `Format As ChartFormat  (read-only)`
- `Height As Double  (read-only)`
- `Width As Double  (read-only)`
- `Top As Double  (read-only)`
- `Left As Double  (read-only)`
- `Name As String  (read-only)`
- `IsTotal As Boolean  (read/write)`

## Methods (9)

- `ClearFormats() As Variant`
- `Copy() As Variant`
- `Delete() As Variant`
- `Paste() As Variant`
- `Select() As Variant`
- `ApplyDataLabels([Type As XlDataLabelsType], [LegendKey As Variant], [AutoText As Variant], [HasLeaderLines As Variant], [ShowSeriesName As Variant], [ShowCategoryName As Variant], [ShowValue As Variant], [ShowPercentage As Variant], [ShowBubbleSize As Variant], [Separator As Variant]) As Variant`
- `PieSliceLocation(loc As XlPieSliceLocation, [Index As XlPieSliceIndex]) As Double`
- `SetProperty(ID As String, Value As Variant)`
- `GetProperty(ID As String) As Variant`
