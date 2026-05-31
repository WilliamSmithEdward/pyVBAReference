# PivotCache

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002441C-0000-0000-C000-000000000046}  

## Properties (34)

- `Application As Application  (read-only)`
- `Creator As XlCreator  (read-only)`
- `Parent As Object  (read-only)`
- `BackgroundQuery As Boolean  (read/write)`
- `Connection As Variant  (read/write)`
- `EnableRefresh As Boolean  (read/write)`
- `Index As Long  (read-only)`
- `MemoryUsed As Long  (read-only)`
- `OptimizeCache As Boolean  (read/write)`
- `RecordCount As Long  (read-only)`
- `RefreshDate As Date  (read-only)`
- `RefreshName As String  (read-only)`
- `RefreshOnFileOpen As Boolean  (read/write)`
- `SavePassword As Boolean  (read/write)`
- `SourceData As Variant  (read/write)`
- `CommandText As Variant  (read/write)`
- `CommandType As XlCmdType  (read/write)`
- `QueryType As XlQueryType  (read-only)`
- `MaintainConnection As Boolean  (read/write)`
- `RefreshPeriod As Long  (read/write)`
- `Recordset As Object  (read/write)`
- `LocalConnection As Variant  (read/write)`
- `UseLocalConnection As Boolean  (read/write)`
- `ADOConnection As Object  (read-only)`
- `IsConnected As Boolean  (read-only)`
- `OLAP As Boolean  (read-only)`
- `SourceType As XlPivotTableSourceType  (read-only)`
- `MissingItemsLimit As XlPivotTableMissingItems  (read/write)`
- `SourceConnectionFile As String  (read/write)`
- `SourceDataFile As String  (read-only)`
- `RobustConnect As XlRobustConnect  (read/write)`
- `WorkbookConnection As WorkbookConnection  (read-only)`
- `Version As XlPivotTableVersionList  (read-only)`
- `UpgradeOnRefresh As Boolean  (read/write)`

## Methods (6)

- `Refresh()`
- `ResetTimer()`
- `CreatePivotTable(TableDestination As Variant, [TableName As Variant], [ReadData As Variant], [DefaultVersion As Variant]) As PivotTable`
- `MakeConnection()`
- `SaveAsODC(ODCFileName As String, [Description As Variant], [Keywords As Variant])`
- `CreatePivotChart(ChartDestination As Variant, [XlChartType As Variant], [Left As Variant], [Top As Variant], [Width As Variant], [Height As Variant]) As Shape`
