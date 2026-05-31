# OLEDBConnection

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002448D-0000-0000-C000-000000000046}  

## Properties (33)

- `Application As Application  (read-only)`
- `Creator As XlCreator  (read-only)`
- `Parent As Object  (read-only)`
- `ADOConnection As Object  (read-only)`
- `BackgroundQuery As Boolean  (read/write)`
- `CommandText As Variant  (read/write)`
- `CommandType As XlCmdType  (read/write)`
- `Connection As Variant  (read/write)`
- `EnableRefresh As Boolean  (read/write)`
- `LocalConnection As Variant  (read/write)`
- `MaintainConnection As Boolean  (read/write)`
- `RefreshDate As Date  (read-only)`
- `Refreshing As Boolean  (read-only)`
- `RefreshOnFileOpen As Boolean  (read/write)`
- `RefreshPeriod As Long  (read/write)`
- `RobustConnect As XlRobustConnect  (read/write)`
- `SavePassword As Boolean  (read/write)`
- `SourceConnectionFile As String  (read/write)`
- `SourceDataFile As String  (read/write)`
- `OLAP As Boolean  (read-only)`
- `UseLocalConnection As Boolean  (read/write)`
- `MaxDrillthroughRecords As Long  (read/write)`
- `IsConnected As Boolean  (read-only)`
- `ServerCredentialsMethod As XlCredentialsMethod  (read/write)`
- `ServerSSOApplicationID As String  (read/write)`
- `AlwaysUseConnectionFile As Boolean  (read/write)`
- `ServerFillColor As Boolean  (read/write)`
- `ServerFontStyle As Boolean  (read/write)`
- `ServerNumberFormat As Boolean  (read/write)`
- `ServerTextColor As Boolean  (read/write)`
- `RetrieveInOfficeUILang As Boolean  (read/write)`
- `CalculatedMembers As CalculatedMembers  (read-only)`
- `LocaleID As Long  (read/write)`

## Methods (5)

- `CancelRefresh()`
- `MakeConnection()`
- `Refresh()`
- `SaveAsODC(ODCFileName As String, [Description As Variant], [Keywords As Variant])`
- `Reconnect()`
