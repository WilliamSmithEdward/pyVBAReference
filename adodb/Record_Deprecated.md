# Record_Deprecated

**Type:** Dispatch Interface  
**Library:** Microsoft ActiveX Data Objects 6.1 Library  
**GUID:** {00000562-0000-0010-8000-00AA006D2EA4}  

## Properties (8)

- `Properties As Properties  (read-only)`
- `ActiveConnection As Variant  (read/write)`
- `State As ObjectStateEnum  (read-only)`
- `Source As Variant  (read/write)`
- `Mode As ConnectModeEnum  (read/write)`
- `ParentURL As String  (read-only)`
- `Fields As Fields_Deprecated  (read-only)`
- `RecordType As RecordTypeEnum  (read-only)`

## Methods (7)

- `MoveRecord([Source As String], [Destination As String], [UserName As String], [Password As String], [Options As MoveRecordOptionsEnum], [Async As Boolean]) As String`
- `CopyRecord([Source As String], [Destination As String], [UserName As String], [Password As String], [Options As CopyRecordOptionsEnum], [Async As Boolean]) As String`
- `DeleteRecord([Source As String], [Async As Boolean])`
- `Open([Source As Variant], [ActiveConnection As Variant], [Mode As ConnectModeEnum], [CreateOptions As RecordCreateOptionsEnum], [Options As RecordOpenOptionsEnum], [UserName As String], [Password As String])`
- `Close()`
- `GetChildren() As _Recordset_Deprecated`
- `Cancel()`
