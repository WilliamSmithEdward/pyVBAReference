# Command

**Type:** Class  
**Library:** Microsoft ActiveX Data Objects 6.1 Library  
**GUID:** {00000507-0000-0010-8000-00AA006D2EA4}  

## Properties (12)

- `Properties As Properties  (read-only)`
- `ActiveConnection As _Connection  (read/write)`
- `CommandText As String  (read/write)`
- `CommandTimeout As Long  (read/write)`
- `Prepared As Boolean  (read/write)`
- `Parameters As Parameters  (read-only)`
- `CommandType As CommandTypeEnum  (read/write)`
- `Name As String  (read/write)`
- `State As Long  (read-only)`
- `CommandStream As Variant  (read/write)`
- `Dialect As String  (read/write)`
- `NamedParameters As Boolean  (read/write)`

## Methods (3)

- `Execute([RecordsAffected As Variant], [Parameters As Variant], [Options As Long]) As _Recordset`
- `CreateParameter([Name As String], [Type As DataTypeEnum], [Direction As ParameterDirectionEnum], [Size As Long], [Value As Variant]) As _Parameter`
- `Cancel()`
