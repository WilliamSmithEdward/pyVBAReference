# TextBox2

**Type:** Class  
**Library:** Microsoft Forms 2.0 Object Library  
**GUID:** {DD4CB8C5-F540-47FF-84D7-67390D2743CA}  

## Properties (39)

- `AutoSize As Boolean  (read/write)`
- `AutoTab As Boolean  (read/write)`
- `AutoWordSelect As Boolean  (read/write)`
- `BackColor As OLE_COLOR  (read/write)`
- `BackStyle As fmBackStyle  (read/write)`
- `BorderColor As OLE_COLOR  (read/write)`
- `BorderStyle As fmBorderStyle  (read/write)`
- `CanPaste As Boolean  (read-only)`
- `CurLine As Long  (read/write)`
- `CurTargetX As Long  (read-only)`
- `CurX As Long  (read/write)`
- `EnterKeyBehavior As Boolean  (read/write)`
- `Enabled As Boolean  (read/write)`
- `Font As Font  (read/write)`
- `ForeColor As OLE_COLOR  (read/write)`
- `HideSelection As Boolean  (read/write)`
- `IntegralHeight As Boolean  (read/write)`
- `LineCount As Long  (read-only)`
- `Locked As Boolean  (read/write)`
- `MaxLength As Long  (read/write)`
- `MouseIcon As Picture  (read/write)`
- `MousePointer As fmMousePointer  (read/write)`
- `MultiLine As Boolean  (read/write)`
- `PasswordChar As String  (read/write)`
- `ScrollBars As fmScrollBars  (read/write)`
- `SelectionMargin As Boolean  (read/write)`
- `SelLength As Long  (read/write)`
- `SelStart As Long  (read/write)`
- `SelText As String  (read/write)`
- `SpecialEffect As fmSpecialEffect  (read/write)`
- `TabKeyBehavior As Boolean  (read/write)`
- `Text As String  (read/write)`
- `TextAlign As fmTextAlign  (read/write)`
- `TextLength As Long  (read-only)`
- `Value As Variant  (read/write)`
- `WordWrap As Boolean  (read/write)`
- `IMEMode As fmIMEMode  (read/write)`
- `EnterFieldBehavior As fmEnterFieldBehavior  (read/write)`
- `DragBehavior As fmDragBehavior  (read/write)`

## Methods (3)

- `Copy()`
- `Cut()`
- `Paste()`

## Events (12)

- `BeforeDragOver(Cancel As ReturnBoolean, Data As DataObject, X As Single, Y As Single, DragState As fmDragState, Effect As ReturnEffect, Shift As Integer)`
- `BeforeDropOrPaste(Cancel As ReturnBoolean, Action As fmAction, Data As DataObject, X As Single, Y As Single, Effect As ReturnEffect, Shift As Integer)`
- `Change()`
- `DblClick(Cancel As ReturnBoolean)`
- `DropButtonClick()`
- `Error(Number As Integer, Description As ReturnString, SCode As Long, Source As String, HelpFile As String, HelpContext As Long, CancelDisplay As ReturnBoolean)`
- `KeyDown(KeyCode As ReturnInteger, Shift As Integer)`
- `KeyPress(KeyAscii As ReturnInteger)`
- `KeyUp(KeyCode As ReturnInteger, Shift As Integer)`
- `MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)`
- `MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)`
- `MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)`
