# RTD

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002446E-0000-0000-C000-000000000046}  

Represents a real-time data object.

**Remarks:** Use the RTD property of the Application object to return an RTD object.

## Properties (1)

- `ThrottleInterval As Long  (read/write)`  
  Returns or sets a Long indicating the time interval between updates. Read/write.

## Methods (2)

- `RefreshData()`  
  Requests an update of real-time data from the real-time data server.
- `RestartServers()`  
  Reconnects to a real-time data server (RTD).
