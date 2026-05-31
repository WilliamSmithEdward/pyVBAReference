# IPivotCache

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002441C-0001-0000-C000-000000000046}  

## Properties (34)

- `Application As HRESULT  (read-only)`
- `Creator As HRESULT  (read-only)`
- `Parent As HRESULT  (read-only)`
- `BackgroundQuery As HRESULT  (read/write)`
- `Connection As HRESULT  (read/write)`
- `EnableRefresh As HRESULT  (read/write)`
- `Index As HRESULT  (read-only)`
- `MemoryUsed As HRESULT  (read-only)`
- `OptimizeCache As HRESULT  (read/write)`
- `RecordCount As HRESULT  (read-only)`
- `RefreshDate As HRESULT  (read-only)`
- `RefreshName As HRESULT  (read-only)`
- `RefreshOnFileOpen As HRESULT  (read/write)`
- `SavePassword As HRESULT  (read/write)`
- `SourceData As HRESULT  (read/write)`
- `CommandText As HRESULT  (read/write)`
- `CommandType As HRESULT  (read/write)`
- `QueryType As HRESULT  (read-only)`
- `MaintainConnection As HRESULT  (read/write)`
- `RefreshPeriod As HRESULT  (read/write)`
- `Recordset As HRESULT  (read/write)`
- `LocalConnection As HRESULT  (read/write)`
- `UseLocalConnection As HRESULT  (read/write)`
- `ADOConnection As HRESULT  (read-only)`
- `IsConnected As HRESULT  (read-only)`
- `OLAP As HRESULT  (read-only)`
- `SourceType As HRESULT  (read-only)`
- `MissingItemsLimit As HRESULT  (read/write)`
- `SourceConnectionFile As HRESULT  (read/write)`
- `SourceDataFile As HRESULT  (read-only)`
- `RobustConnect As HRESULT  (read/write)`
- `WorkbookConnection As HRESULT  (read-only)`
- `Version As HRESULT  (read-only)`
- `UpgradeOnRefresh As HRESULT  (read/write)`

## Methods (6)

- `Refresh()`
- `ResetTimer()`
- `CreatePivotTable(TableDestination As Variant, [TableName As Variant], [ReadData As Variant], [DefaultVersion As Variant], RHS As PivotTable)`
- `MakeConnection()`
- `SaveAsODC(ODCFileName As String, [Description As Variant], [Keywords As Variant])`
- `CreatePivotChart(ChartDestination As Variant, [XlChartType As Variant], [Left As Variant], [Top As Variant], [Width As Variant], [Height As Variant], RHS As Shape)`
