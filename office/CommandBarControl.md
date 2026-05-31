# CommandBarControl

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0308-0000-0000-C000-000000000046}  

## Properties (25)

- `Application As Object  (read-only)`
- `Creator As Long  (read-only)`
- `BeginGroup As Boolean  (read/write)`
- `BuiltIn As Boolean  (read-only)`
- `Caption As String  (read/write)`
- `DescriptionText As String  (read/write)`
- `Enabled As Boolean  (read/write)`
- `Height As Long  (read/write)`
- `HelpContextId As Long  (read/write)`
- `HelpFile As String  (read/write)`
- `Id As Long  (read-only)`
- `Index As Long  (read-only)`
- `Left As Long  (read-only)`
- `OLEUsage As MsoControlOLEUsage  (read/write)`
- `OnAction As String  (read/write)`
- `Parent As CommandBar  (read-only)`
- `Parameter As String  (read/write)`
- `Priority As Long  (read/write)`
- `Tag As String  (read/write)`
- `TooltipText As String  (read/write)`
- `Top As Long  (read-only)`
- `Type As MsoControlType  (read-only)`
- `Visible As Boolean  (read/write)`
- `Width As Long  (read/write)`
- `IsPriorityDropped As Boolean  (read-only)`

## Methods (6)

- `Copy([Bar As Variant], [Before As Variant]) As CommandBarControl`
- `Delete([Temporary As Variant])`
- `Execute()`
- `Move([Bar As Variant], [Before As Variant]) As CommandBarControl`
- `Reset()`
- `SetFocus()`
