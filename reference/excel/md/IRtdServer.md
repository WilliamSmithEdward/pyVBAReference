# IRtdServer

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {EC0E6191-DB51-11D3-8F3E-00C04F3651B8}  

Represents an interface for a real-time data server.

**Remarks:** The IRTDServer object can be instantiated or created only by implementing the IRTDServer interface by using the Implements keyword.

## Methods (6)

- `ServerStart(CallbackObject As IRTDUpdateEvent) As Long`  
  The ServerStart method is called immediately after a real-time data (RTD) server is instantiated. Returns a Long. A negative value or zero indicates failure to start the server; a positive value indicates success.
    - `CallbackObject As IRTDUpdateEvent` (required): The callback object.
- `ConnectData(TopicID As Long, Strings As SAFEARRAY(Variant), GetNewValues As Boolean) As Variant`  
  Adds new topics from a real-time data (RTD) server. The ConnectData method is called when a file is opened that contains real-time data functions or when a user types in a new formula that contains the RTD function.
    - `TopicID As Long` (required): A unique value, assigned by Microsoft Excel, that identifies the topic.
    - `GetNewValues As Boolean` (required): True to determine if new values are to be acquired.
- `RefreshData(TopicCount As Long) As SAFEARRAY(Variant)`  
  This method is called by Microsoft Excel to get new data. Returns a Variant.
    - `TopicCount As Long` (required): The real-time data (RTD) server must change the value of the TopicCount to the number of elements in the array returned.
- `DisconnectData(TopicID As Long)`  
  Notifies a real-time data (RTD) server application that a topic is no longer in use.
    - `TopicID As Long` (required): A unique value assigned to the topic assigned by Microsoft Excel.
- `Heartbeat() As Long`  
  Determines if the real-time data (RTD) server is still active. Returns a Long value. Zero or a negative number indicates failure; a positive number indicates that the server is active.
- `ServerTerminate()`  
  Terminates the connection to the real-time data (RTD) server.
