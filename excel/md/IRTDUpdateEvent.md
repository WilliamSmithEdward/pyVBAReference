# IRTDUpdateEvent

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {A43788C1-D91B-11D3-8F39-00C04F3651B8}  

Represents real-time data update events.

**Remarks:** To instantiate or return an IRTDUpdateEvent object, declare a variable as an IRTDUpdateEvent object, and then use that variable as a callback object.

## Properties (1)

- `HeartbeatInterval As Long  (read/write)`  
  Returns or sets a Long for the interval between updates for real-time data. Read/write.

## Methods (2)

- `UpdateNotify()`  
  The real-time data (RTD) server uses this method to notify Microsoft Excel that new data has been received.
- `Disconnect()`  
  Instructs the real-time data (RTD) server to disconnect from the specified IRTDUpdateEvent object.
