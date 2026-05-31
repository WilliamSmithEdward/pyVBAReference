# ODBCConnection

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002448E-0000-0000-C000-000000000046}  

## Properties (20)

- `Application As Application  (read-only)`
- `Creator As XlCreator  (read-only)`
- `Parent As Object  (read-only)`
- `BackgroundQuery As Boolean  (read/write)`
- `CommandText As Variant  (read/write)`
- `CommandType As XlCmdType  (read/write)`
- `Connection As Variant  (read/write)`
- `EnableRefresh As Boolean  (read/write)`
- `RefreshDate As Date  (read-only)`
- `Refreshing As Boolean  (read-only)`
- `RefreshOnFileOpen As Boolean  (read/write)`
- `RefreshPeriod As Long  (read/write)`
- `RobustConnect As XlRobustConnect  (read/write)`
- `SavePassword As Boolean  (read/write)`
- `SourceConnectionFile As String  (read/write)`
- `SourceData As Variant  (read/write)`
- `SourceDataFile As String  (read/write)`
- `ServerCredentialsMethod As XlCredentialsMethod  (read/write)`
- `ServerSSOApplicationID As String  (read/write)`
- `AlwaysUseConnectionFile As Boolean  (read/write)`

## Methods (3)

- `CancelRefresh()`
- `Refresh()`
- `SaveAsODC(ODCFileName As String, [Description As Variant], [Keywords As Variant])`
