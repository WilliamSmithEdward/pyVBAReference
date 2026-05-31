# ComboBox

**Type:** Class  
**Library:** Microsoft Forms 2.0 Object Library  
**GUID:** {8BD21D30-EC42-11CE-9E0D-00AA006002F3}  

## Properties (50)

- `AutoSize As Boolean  (read/write)`
- `AutoTab As Boolean  (read/write)`
- `AutoWordSelect As Boolean  (read/write)`
- `BackColor As OLE_COLOR  (read/write)`
- `BackStyle As fmBackStyle  (read/write)`
- `BorderColor As OLE_COLOR  (read/write)`
- `BorderStyle As fmBorderStyle  (read/write)`
- `BoundColumn As Variant  (read/write)`
- `CanPaste As Boolean  (read-only)`
- `ColumnCount As Long  (read/write)`
- `ColumnHeads As Boolean  (read/write)`
- `ColumnWidths As String  (read/write)`
- `CurTargetX As Long  (read-only)`
- `CurX As Long  (read/write)`
- `DropButtonStyle As fmDropButtonStyle  (read/write)`
- `Enabled As Boolean  (read/write)`
- `Font As Font  (read/write)`
- `ForeColor As OLE_COLOR  (read/write)`
- `HideSelection As Boolean  (read/write)`
- `LineCount As Long  (read-only)`
- `ListCount As Long  (read-only)`
- `ListIndex As Variant  (read/write)`
- `ListRows As Long  (read/write)`
- `ListStyle As fmListStyle  (read/write)`
- `ListWidth As Variant  (read/write)`
- `Locked As Boolean  (read/write)`
- `MatchEntry As fmMatchEntry  (read/write)`
- `MatchFound As Boolean  (read-only)`
- `MatchRequired As Boolean  (read/write)`
- `MaxLength As Long  (read/write)`
- `MouseIcon As Picture  (read/write)`
- `MousePointer As fmMousePointer  (read/write)`
- `SelectionMargin As Boolean  (read/write)`
- `SelLength As Long  (read/write)`
- `SelStart As Long  (read/write)`
- `SelText As String  (read/write)`
- `ShowDropButtonWhen As fmShowDropButtonWhen  (read/write)`
- `SpecialEffect As fmSpecialEffect  (read/write)`
- `Style As fmStyle  (read/write)`
- `Text As String  (read/write)`
- `TextAlign As fmTextAlign  (read/write)`
- `TextColumn As Variant  (read/write)`
- `TextLength As Long  (read-only)`
- `TopIndex As Variant  (read/write)`
- `Value As Variant  (read/write)`
- `Column As Variant  (read/write)`
- `List As Variant  (read/write)`
- `IMEMode As fmIMEMode  (read/write)`
- `EnterFieldBehavior As fmEnterFieldBehavior  (read/write)`
- `DragBehavior As fmDragBehavior  (read/write)`

## Methods (7)

- `AddItem([pvargItem As Variant], [pvargIndex As Variant])`
- `Clear()`
- `DropDown()`
- `RemoveItem(pvargIndex As Variant)`
- `Copy()`
- `Cut()`
- `Paste()`

## Events (13)

- `BeforeDragOver(Cancel As ReturnBoolean, Data As DataObject, X As Single, Y As Single, DragState As fmDragState, Effect As ReturnEffect, Shift As Integer)`
- `BeforeDropOrPaste(Cancel As ReturnBoolean, Action As fmAction, Data As DataObject, X As Single, Y As Single, Effect As ReturnEffect, Shift As Integer)`
- `Change()`
- `Click()`
- `DblClick(Cancel As ReturnBoolean)`
- `DropButtonClick()`
- `Error(Number As Integer, Description As ReturnString, SCode As Long, Source As String, HelpFile As String, HelpContext As Long, CancelDisplay As ReturnBoolean)`
- `KeyDown(KeyCode As ReturnInteger, Shift As Integer)`
- `KeyPress(KeyAscii As ReturnInteger)`
- `KeyUp(KeyCode As ReturnInteger, Shift As Integer)`
- `MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)`
- `MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)`
- `MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)`
