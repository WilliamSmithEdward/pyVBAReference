# PickerDialog

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03E6-0000-0000-C000-000000000046}  

## Properties (5)

- `Application As Object  (read-only)`
- `Creator As Long  (read-only)`
- `DataHandlerId As String  (read/write)`
- `Title As String  (read/write)`
- `Properties As PickerProperties  (read-only)`

## Methods (3)

- `CreatePickerResults() As PickerResults`
- `Show([IsMultiSelect As Boolean], [ExistingResults As PickerResults]) As PickerResults`
- `Resolve(TokenText As String, duplicateDlgMode As Long) As PickerResults`
