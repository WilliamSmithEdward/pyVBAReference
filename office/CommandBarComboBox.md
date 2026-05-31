# CommandBarComboBox

**Type:** Class  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {55F88897-7708-11D1-ACEB-006008961DA5}  

## Properties (33)

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
- `DropDownLines As Long  (read/write)`
- `DropDownWidth As Long  (read/write)`
- `List As String  (read/write)`
- `ListCount As Long  (read-only)`
- `ListHeaderCount As Long  (read/write)`
- `ListIndex As Long  (read/write)`
- `Style As MsoComboStyle  (read/write)`
- `Text As String  (read/write)`

## Methods (9)

- `Copy([Bar As Variant], [Before As Variant]) As CommandBarControl`
- `Delete([Temporary As Variant])`
- `Execute()`
- `Move([Bar As Variant], [Before As Variant]) As CommandBarControl`
- `Reset()`
- `SetFocus()`
- `AddItem(Text As String, [Index As Variant])`
- `Clear()`
- `RemoveItem(Index As Long)`

## Events (1)

- `Change(Ctrl As CommandBarComboBox)`
