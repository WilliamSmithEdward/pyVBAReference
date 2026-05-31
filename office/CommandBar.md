# CommandBar

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0304-0000-0000-C000-000000000046}  

## Properties (20)

- `Application As Object  (read-only)`
- `Creator As Long  (read-only)`
- `BuiltIn As Boolean  (read-only)`
- `Context As String  (read/write)`
- `Controls As CommandBarControls  (read-only)`
- `Enabled As Boolean  (read/write)`
- `Height As Long  (read/write)`
- `Index As Long  (read-only)`
- `Left As Long  (read/write)`
- `Name As String  (read/write)`
- `NameLocal As String  (read/write)`
- `Parent As Object  (read-only)`
- `Position As MsoBarPosition  (read/write)`
- `RowIndex As Long  (read/write)`
- `Protection As MsoBarProtection  (read/write)`
- `Top As Long  (read/write)`
- `Type As MsoBarType  (read-only)`
- `Visible As Boolean  (read/write)`
- `Width As Long  (read/write)`
- `AdaptiveMenu As Boolean  (read/write)`

## Methods (4)

- `Delete()`
- `FindControl([Type As Variant], [Id As Variant], [Tag As Variant], [Visible As Variant], [Recursive As Variant]) As CommandBarControl`
- `Reset()`
- `ShowPopup([x As Variant], [y As Variant])`
