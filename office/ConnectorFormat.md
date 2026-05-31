# ConnectorFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0313-0000-0000-C000-000000000046}  

## Properties (10)

- `Application As Object  (read-only)`
- `Creator As Long  (read-only)`
- `Parent As Object  (read-only)`
- `BeginConnected As MsoTriState  (read-only)`
- `BeginConnectedShape As Shape  (read-only)`
- `BeginConnectionSite As Long  (read-only)`
- `EndConnected As MsoTriState  (read-only)`
- `EndConnectedShape As Shape  (read-only)`
- `EndConnectionSite As Long  (read-only)`
- `Type As MsoConnectorType  (read/write)`

## Methods (4)

- `BeginConnect(ConnectedShape As Shape, ConnectionSite As Long)`
- `BeginDisconnect()`
- `EndConnect(ConnectedShape As Shape, ConnectionSite As Long)`
- `EndDisconnect()`
