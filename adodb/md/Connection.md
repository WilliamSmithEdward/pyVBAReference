# Connection

**Type:** Class  
**Library:** Microsoft ActiveX Data Objects 6.1 Library  
**GUID:** {00000514-0000-0010-8000-00AA006D2EA4}  

## Properties (13)

- `Properties As Properties  (read-only)`
- `ConnectionString As String  (read/write)`
- `CommandTimeout As Long  (read/write)`
- `ConnectionTimeout As Long  (read/write)`
- `Version As String  (read-only)`
- `Errors As Errors  (read-only)`
- `DefaultDatabase As String  (read/write)`
- `IsolationLevel As IsolationLevelEnum  (read/write)`
- `Attributes As Long  (read/write)`
- `CursorLocation As CursorLocationEnum  (read/write)`
- `Mode As ConnectModeEnum  (read/write)`
- `Provider As String  (read/write)`
- `State As Long  (read-only)`

## Methods (8)

- `Close()`
- `Execute(CommandText As String, [RecordsAffected As Variant], [Options As Long]) As _Recordset`
- `BeginTrans() As Long`
- `CommitTrans()`
- `RollbackTrans()`
- `Open([ConnectionString As String], [UserID As String], [Password As String], [Options As Long])`
- `OpenSchema(Schema As SchemaEnum, [Restrictions As Variant], [SchemaID As Variant]) As _Recordset`
- `Cancel()`

## Events (9)

- `InfoMessage(pError As Error, adStatus As EventStatusEnum, pConnection As _Connection)`
- `BeginTransComplete(TransactionLevel As Long, pError As Error, adStatus As EventStatusEnum, pConnection As _Connection)`
- `CommitTransComplete(pError As Error, adStatus As EventStatusEnum, pConnection As _Connection)`
- `RollbackTransComplete(pError As Error, adStatus As EventStatusEnum, pConnection As _Connection)`
- `WillExecute(Source As String, CursorType As CursorTypeEnum, LockType As LockTypeEnum, Options As Long, adStatus As EventStatusEnum, pCommand As _Command, pRecordset As _Recordset, pConnection As _Connection)`
- `ExecuteComplete(RecordsAffected As Long, pError As Error, adStatus As EventStatusEnum, pCommand As _Command, pRecordset As _Recordset, pConnection As _Connection)`
- `WillConnect(ConnectionString As String, UserID As String, Password As String, Options As Long, adStatus As EventStatusEnum, pConnection As _Connection)`
- `ConnectComplete(pError As Error, adStatus As EventStatusEnum, pConnection As _Connection)`
- `Disconnect(adStatus As EventStatusEnum, pConnection As _Connection)`
