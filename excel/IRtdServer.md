# IRtdServer

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {EC0E6191-DB51-11D3-8F3E-00C04F3651B8}  

## Methods (6)

- `ServerStart(CallbackObject As IRTDUpdateEvent) As Long`
- `ConnectData(TopicID As Long, Strings As SAFEARRAY(Variant), GetNewValues As Boolean) As Variant`
- `RefreshData(TopicCount As Long) As SAFEARRAY(Variant)`
- `DisconnectData(TopicID As Long)`
- `Heartbeat() As Long`
- `ServerTerminate()`
