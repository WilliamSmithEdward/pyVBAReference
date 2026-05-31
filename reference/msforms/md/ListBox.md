# ListBox

**Type:** Class  
**Library:** Microsoft Forms 2.0 Object Library  
**GUID:** {8BD21D20-EC42-11CE-9E0D-00AA006002F3}  

## Properties (29)

- `BackColor As OLE_COLOR  (read/write)`
- `BorderColor As OLE_COLOR  (read/write)`
- `BorderStyle As fmBorderStyle  (read/write)`
- `BoundColumn As Variant  (read/write)`
- `ColumnCount As Long  (read/write)`
- `ColumnHeads As Boolean  (read/write)`
- `ColumnWidths As String  (read/write)`
- `Enabled As Boolean  (read/write)`
- `Font As Font  (read/write)`
- `ForeColor As OLE_COLOR  (read/write)`
- `IntegralHeight As Boolean  (read/write)`
- `ListCount As Long  (read-only)`
- `ListIndex As Variant  (read/write)`
- `ListStyle As fmListStyle  (read/write)`
- `Locked As Boolean  (read/write)`
- `MatchEntry As fmMatchEntry  (read/write)`
- `MouseIcon As Picture  (read/write)`
- `MousePointer As fmMousePointer  (read/write)`
- `MultiSelect As fmMultiSelect  (read/write)`
- `SpecialEffect As fmSpecialEffect  (read/write)`
- `Text As String  (read/write)`
- `TextColumn As Variant  (read/write)`
- `TopIndex As Variant  (read/write)`
- `Value As Variant  (read/write)`
- `Column As Variant  (read/write)`
- `List As Variant  (read/write)`
- `Selected As Boolean  (read/write)`
- `IMEMode As fmIMEMode  (read/write)`
- `TextAlign As fmTextAlign  (read/write)`

## Methods (3)

- `AddItem([pvargItem As Variant], [pvargIndex As Variant])`
- `Clear()`
- `RemoveItem(pvargIndex As Variant)`

## Events (12)

- `BeforeDragOver(Cancel As ReturnBoolean, Data As DataObject, X As Single, Y As Single, DragState As fmDragState, Effect As ReturnEffect, Shift As Integer)`
- `BeforeDropOrPaste(Cancel As ReturnBoolean, Action As fmAction, Data As DataObject, X As Single, Y As Single, Effect As ReturnEffect, Shift As Integer)`
- `Change()`
- `Click()`
- `DblClick(Cancel As ReturnBoolean)`
- `Error(Number As Integer, Description As ReturnString, SCode As Long, Source As String, HelpFile As String, HelpContext As Long, CancelDisplay As ReturnBoolean)`
- `KeyDown(KeyCode As ReturnInteger, Shift As Integer)`
- `KeyPress(KeyAscii As ReturnInteger)`
- `KeyUp(KeyCode As ReturnInteger, Shift As Integer)`
- `MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)`
- `MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)`
- `MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)`
