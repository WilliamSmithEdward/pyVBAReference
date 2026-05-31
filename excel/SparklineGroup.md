# SparklineGroup

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244B7-0000-0000-C000-000000000046}  

## Properties (17)

- `Application As Application  (read-only)`
- `Creator As XlCreator  (read-only)`
- `Parent As Object  (read-only)`
- `Count As Long  (read-only)`
- `Item As Sparkline  (read-only)`
- `_NewEnum As IUnknown  (read-only)`
- `Location As Range  (read/write)`
- `SourceData As String  (read/write)`
- `DateRange As String  (read/write)`
- `Type As XlSparkType  (read/write)`
- `SeriesColor As FormatColor  (read-only)`
- `Points As SparkPoints  (read-only)`
- `Axes As SparkAxes  (read-only)`
- `DisplayBlanksAs As XlDisplayBlanksAs  (read/write)`
- `DisplayHidden As Boolean  (read/write)`
- `LineWeight As Variant  (read/write)`
- `PlotBy As XlSparklineRowCol  (read/write)`

## Methods (5)

- `ModifyLocation(Location As Range)`
- `ModifySourceData(SourceData As String)`
- `Modify(Location As Range, SourceData As String)`
- `ModifyDateRange(DateRange As String)`
- `Delete()`
