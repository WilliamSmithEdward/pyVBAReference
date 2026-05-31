# IDataFeedConnection

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244D4-0001-0000-C000-000000000046}  

## Properties (16)

- `Application As HRESULT  (read-only)`
- `Creator As HRESULT  (read-only)`
- `Parent As HRESULT  (read-only)`
- `AlwaysUseConnectionFile As HRESULT  (read/write)`
- `CommandText As HRESULT  (read/write)`
- `CommandType As HRESULT  (read/write)`
- `Connection As HRESULT  (read/write)`
- `EnableRefresh As HRESULT  (read/write)`
- `RefreshDate As HRESULT  (read-only)`
- `Refreshing As HRESULT  (read-only)`
- `RefreshOnFileOpen As HRESULT  (read/write)`
- `RefreshPeriod As HRESULT  (read/write)`
- `SavePassword As HRESULT  (read/write)`
- `ServerCredentialsMethod As HRESULT  (read/write)`
- `SourceConnectionFile As HRESULT  (read/write)`
- `SourceDataFile As HRESULT  (read/write)`

## Methods (3)

- `CancelRefresh()`
- `Refresh()`
- `SaveAsODC(ODCFileName As String, [Description As Variant], [Keywords As Variant])`
