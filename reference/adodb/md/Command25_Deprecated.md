# Command25_Deprecated

**Type:** Dispatch Interface  
**Library:** Microsoft ActiveX Data Objects 6.1 Library  
**GUID:** {0000054E-0000-0010-8000-00AA006D2EA4}  

## Properties (9)

- `Properties As Properties  (read-only)`
- `ActiveConnection As _Connection_Deprecated  (read/write)`
- `CommandText As String  (read/write)`
- `CommandTimeout As Long  (read/write)`
- `Prepared As Boolean  (read/write)`
- `Parameters As Parameters_Deprecated  (read-only)`
- `CommandType As CommandTypeEnum  (read/write)`
- `Name As String  (read/write)`
- `State As Long  (read-only)`

## Methods (3)

- `Execute([RecordsAffected As Variant], [Parameters As Variant], [Options As Long]) As _Recordset_Deprecated`
- `CreateParameter([Name As String], [Type As DataTypeEnum], [Direction As ParameterDirectionEnum], [Size As ADO_LONGPTR], [Value As Variant]) As _Parameter_Deprecated`
- `Cancel()`
