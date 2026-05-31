# IPoint

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002086A-0001-0000-C000-000000000046}  

## Properties (28)

- `Application As HRESULT  (read-only)`
- `Creator As HRESULT  (read-only)`
- `Parent As HRESULT  (read-only)`
- `DataLabel As HRESULT  (read-only)`
- `Explosion As HRESULT  (read/write)`
- `HasDataLabel As HRESULT  (read/write)`
- `InvertIfNegative As HRESULT  (read/write)`
- `MarkerBackgroundColor As HRESULT  (read/write)`
- `MarkerBackgroundColorIndex As HRESULT  (read/write)`
- `MarkerForegroundColor As HRESULT  (read/write)`
- `MarkerForegroundColorIndex As HRESULT  (read/write)`
- `MarkerSize As HRESULT  (read/write)`
- `MarkerStyle As HRESULT  (read/write)`
- `PictureType As HRESULT  (read/write)`
- `ApplyPictToSides As HRESULT  (read/write)`
- `ApplyPictToFront As HRESULT  (read/write)`
- `ApplyPictToEnd As HRESULT  (read/write)`
- `Shadow As HRESULT  (read/write)`
- `SecondaryPlot As HRESULT  (read/write)`
- `Has3DEffect As HRESULT  (read/write)`
- `PictureUnit2 As HRESULT  (read/write)`
- `Format As HRESULT  (read-only)`
- `Height As HRESULT  (read-only)`
- `Width As HRESULT  (read-only)`
- `Top As HRESULT  (read-only)`
- `Left As HRESULT  (read-only)`
- `Name As HRESULT  (read-only)`
- `IsTotal As HRESULT  (read/write)`

## Methods (9)

- `ClearFormats(RHS As Variant)`
- `Copy(RHS As Variant)`
- `Delete(RHS As Variant)`
- `Paste(RHS As Variant)`
- `Select(RHS As Variant)`
- `ApplyDataLabels([Type As XlDataLabelsType], [LegendKey As Variant], [AutoText As Variant], [HasLeaderLines As Variant], [ShowSeriesName As Variant], [ShowCategoryName As Variant], [ShowValue As Variant], [ShowPercentage As Variant], [ShowBubbleSize As Variant], [Separator As Variant], RHS As Variant)`
- `PieSliceLocation(loc As XlPieSliceLocation, [Index As XlPieSliceIndex], RHS As Double)`
- `SetProperty(ID As String, Value As Variant)`
- `GetProperty(ID As String, RHS As Variant)`
