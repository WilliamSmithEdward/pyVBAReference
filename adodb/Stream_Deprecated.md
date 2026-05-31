# Stream_Deprecated

**Type:** Dispatch Interface  
**Library:** Microsoft ActiveX Data Objects 6.1 Library  
**GUID:** {00000565-0000-0010-8000-00AA006D2EA4}  

## Properties (8)

- `Size As ADO_LONGPTR  (read-only)`
- `EOS As Boolean  (read-only)`
- `Position As ADO_LONGPTR  (read/write)`
- `Type As StreamTypeEnum  (read/write)`
- `LineSeparator As LineSeparatorEnum  (read/write)`
- `State As ObjectStateEnum  (read-only)`
- `Mode As ConnectModeEnum  (read/write)`
- `Charset As String  (read/write)`

## Methods (13)

- `Read([NumBytes As Long]) As Variant`
- `Open([Source As Variant], [Mode As ConnectModeEnum], [Options As StreamOpenOptionsEnum], [UserName As String], [Password As String])`
- `Close()`
- `SkipLine()`
- `Write(Buffer As Variant)`
- `SetEOS()`
- `CopyTo(DestStream As _Stream_Deprecated, [CharNumber As ADO_LONGPTR])`
- `Flush()`
- `SaveToFile(FileName As String, [Options As SaveOptionsEnum])`
- `LoadFromFile(FileName As String)`
- `ReadText([NumChars As Long]) As String`
- `WriteText(Data As String, [Options As StreamWriteEnum])`
- `Cancel()`
