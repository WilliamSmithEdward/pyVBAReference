# ConnectionEvents

**Type:** Dispatch Interface  
**Library:** Microsoft ActiveX Data Objects 6.1 Library  
**GUID:** {00001400-0000-0010-8000-00AA006D2EA4}  

## Methods (9)

- `InfoMessage(pError As Error, adStatus As EventStatusEnum, pConnection As _Connection)`
- `BeginTransComplete(TransactionLevel As Long, pError As Error, adStatus As EventStatusEnum, pConnection As _Connection)`
- `CommitTransComplete(pError As Error, adStatus As EventStatusEnum, pConnection As _Connection)`
- `RollbackTransComplete(pError As Error, adStatus As EventStatusEnum, pConnection As _Connection)`
- `WillExecute(Source As String, CursorType As CursorTypeEnum, LockType As LockTypeEnum, Options As Long, adStatus As EventStatusEnum, pCommand As _Command, pRecordset As _Recordset, pConnection As _Connection)`
- `ExecuteComplete(RecordsAffected As Long, pError As Error, adStatus As EventStatusEnum, pCommand As _Command, pRecordset As _Recordset, pConnection As _Connection)`
- `WillConnect(ConnectionString As String, UserID As String, Password As String, Options As Long, adStatus As EventStatusEnum, pConnection As _Connection)`
- `ConnectComplete(pError As Error, adStatus As EventStatusEnum, pConnection As _Connection)`
- `Disconnect(adStatus As EventStatusEnum, pConnection As _Connection)`
