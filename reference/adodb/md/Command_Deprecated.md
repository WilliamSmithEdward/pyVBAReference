# Command_Deprecated

**Type:** Dispatch Interface  
**Library:** Microsoft ActiveX Data Objects 6.1 Library  
**GUID:** {B08400BD-F9D1-4D02-B856-71D5DBA123E9}  

## Properties (12)

- `Properties As Properties  (read-only)`
- `ActiveConnection As _Connection_Deprecated  (read/write)`
- `CommandText As String  (read/write)`
- `CommandTimeout As Long  (read/write)`
- `Prepared As Boolean  (read/write)`
- `Parameters As Parameters_Deprecated  (read-only)`
- `CommandType As CommandTypeEnum  (read/write)`
- `Name As String  (read/write)`
- `State As Long  (read-only)`
- `CommandStream As Variant  (read/write)`
- `Dialect As String  (read/write)`
- `NamedParameters As Boolean  (read/write)`

## Methods (3)

- `Execute([RecordsAffected As Variant], [Parameters As Variant], [Options As Long]) As _Recordset_Deprecated`
- `CreateParameter([Name As String], [Type As DataTypeEnum], [Direction As ParameterDirectionEnum], [Size As ADO_LONGPTR], [Value As Variant]) As _Parameter_Deprecated`
- `Cancel()`
