# IOLEDBConnection

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002448D-0001-0000-C000-000000000046}  

## Properties (33)

- `Application As HRESULT  (read-only)`
- `Creator As HRESULT  (read-only)`
- `Parent As HRESULT  (read-only)`
- `ADOConnection As HRESULT  (read-only)`
- `BackgroundQuery As HRESULT  (read/write)`
- `CommandText As HRESULT  (read/write)`
- `CommandType As HRESULT  (read/write)`
- `Connection As HRESULT  (read/write)`
- `EnableRefresh As HRESULT  (read/write)`
- `LocalConnection As HRESULT  (read/write)`
- `MaintainConnection As HRESULT  (read/write)`
- `RefreshDate As HRESULT  (read-only)`
- `Refreshing As HRESULT  (read-only)`
- `RefreshOnFileOpen As HRESULT  (read/write)`
- `RefreshPeriod As HRESULT  (read/write)`
- `RobustConnect As HRESULT  (read/write)`
- `SavePassword As HRESULT  (read/write)`
- `SourceConnectionFile As HRESULT  (read/write)`
- `SourceDataFile As HRESULT  (read/write)`
- `OLAP As HRESULT  (read-only)`
- `UseLocalConnection As HRESULT  (read/write)`
- `MaxDrillthroughRecords As HRESULT  (read/write)`
- `IsConnected As HRESULT  (read-only)`
- `ServerCredentialsMethod As HRESULT  (read/write)`
- `ServerSSOApplicationID As HRESULT  (read/write)`
- `AlwaysUseConnectionFile As HRESULT  (read/write)`
- `ServerFillColor As HRESULT  (read/write)`
- `ServerFontStyle As HRESULT  (read/write)`
- `ServerNumberFormat As HRESULT  (read/write)`
- `ServerTextColor As HRESULT  (read/write)`
- `RetrieveInOfficeUILang As HRESULT  (read/write)`
- `CalculatedMembers As HRESULT  (read-only)`
- `LocaleID As HRESULT  (read/write)`

## Methods (5)

- `CancelRefresh()`
- `MakeConnection()`
- `Refresh()`
- `SaveAsODC(ODCFileName As String, [Description As Variant], [Keywords As Variant])`
- `Reconnect()`
