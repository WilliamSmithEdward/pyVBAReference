# DataFeedConnection

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244D4-0000-0000-C000-000000000046}  

## Properties (16)

- `Application As Application  (read-only)`
- `Creator As XlCreator  (read-only)`
- `Parent As Object  (read-only)`
- `AlwaysUseConnectionFile As Boolean  (read/write)`
- `CommandText As Variant  (read/write)`
- `CommandType As XlCmdType  (read/write)`
- `Connection As Variant  (read/write)`
- `EnableRefresh As Boolean  (read/write)`
- `RefreshDate As Date  (read-only)`
- `Refreshing As Boolean  (read-only)`
- `RefreshOnFileOpen As Boolean  (read/write)`
- `RefreshPeriod As Long  (read/write)`
- `SavePassword As Boolean  (read/write)`
- `ServerCredentialsMethod As XlCredentialsMethod  (read/write)`
- `SourceConnectionFile As String  (read/write)`
- `SourceDataFile As String  (read/write)`

## Methods (3)

- `CancelRefresh()`
- `Refresh()`
- `SaveAsODC(ODCFileName As String, [Description As Variant], [Keywords As Variant])`
