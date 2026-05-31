# IConnectorFormat

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002443E-0001-0000-C000-000000000046}  

## Properties (10)

- `Application As HRESULT  (read-only)`
- `Creator As HRESULT  (read-only)`
- `Parent As HRESULT  (read-only)`
- `BeginConnected As HRESULT  (read-only)`
- `BeginConnectedShape As HRESULT  (read-only)`
- `BeginConnectionSite As HRESULT  (read-only)`
- `EndConnected As HRESULT  (read-only)`
- `EndConnectedShape As HRESULT  (read-only)`
- `EndConnectionSite As HRESULT  (read-only)`
- `Type As HRESULT  (read/write)`

## Methods (4)

- `BeginConnect(ConnectedShape As Shape, ConnectionSite As Long)`
- `BeginDisconnect()`
- `EndConnect(ConnectedShape As Shape, ConnectionSite As Long)`
- `EndDisconnect()`
