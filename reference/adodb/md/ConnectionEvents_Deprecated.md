# ConnectionEvents_Deprecated

**Type:** Dispatch Interface  
**Library:** Microsoft ActiveX Data Objects 6.1 Library  
**GUID:** {00000400-0000-0010-8000-00AA006D2EA4}  

## Methods (9)

- `InfoMessage(pError As Error, adStatus As EventStatusEnum, pConnection As _Connection_Deprecated)`
- `BeginTransComplete(TransactionLevel As Long, pError As Error, adStatus As EventStatusEnum, pConnection As _Connection_Deprecated)`
- `CommitTransComplete(pError As Error, adStatus As EventStatusEnum, pConnection As _Connection_Deprecated)`
- `RollbackTransComplete(pError As Error, adStatus As EventStatusEnum, pConnection As _Connection_Deprecated)`
- `WillExecute(Source As String, CursorType As CursorTypeEnum, LockType As LockTypeEnum, Options As Long, adStatus As EventStatusEnum, pCommand As _Command_Deprecated, pRecordset As _Recordset_Deprecated, pConnection As _Connection_Deprecated)`
- `ExecuteComplete(RecordsAffected As Long, pError As Error, adStatus As EventStatusEnum, pCommand As _Command_Deprecated, pRecordset As _Recordset_Deprecated, pConnection As _Connection_Deprecated)`
- `WillConnect(ConnectionString As String, UserID As String, Password As String, Options As Long, adStatus As EventStatusEnum, pConnection As _Connection_Deprecated)`
- `ConnectComplete(pError As Error, adStatus As EventStatusEnum, pConnection As _Connection_Deprecated)`
- `Disconnect(adStatus As EventStatusEnum, pConnection As _Connection_Deprecated)`
