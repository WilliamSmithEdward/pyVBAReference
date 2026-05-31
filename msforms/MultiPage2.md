# MultiPage2

**Type:** Class  
**Library:** Microsoft Forms 2.0 Object Library  
**GUID:** {6C1B3099-127A-4BE1-93BC-DD4771EEEF90}  

## Properties (12)

- `BackColor As OLE_COLOR  (read/write)`
- `ForeColor As OLE_COLOR  (read/write)`
- `Font As Font  (read/write)`
- `MultiRow As Boolean  (read/write)`
- `Style As fmTabStyle  (read/write)`
- `TabOrientation As fmTabOrientation  (read/write)`
- `Enabled As Boolean  (read/write)`
- `SelectedItem As Page  (read-only)`
- `Pages As Pages  (read-only)`
- `Value As Long  (read/write)`
- `TabFixedWidth As Single  (read/write)`
- `TabFixedHeight As Single  (read/write)`

## Methods (4)

- `_SetTabFixedWidth(Width As Long)`
- `_GetTabFixedWidth(Width As Long)`
- `_SetTabFixedHeight(Height As Long)`
- `_GetTabFixedHeight(Height As Long)`

## Events (17)

- `AddControl(Index As Long, Control As Control)`
- `BeforeDragOver(Index As Long, Cancel As ReturnBoolean, Control As Control, Data As DataObject, X As Single, Y As Single, State As fmDragState, Effect As ReturnEffect, Shift As Integer)`
- `BeforeDropOrPaste(Index As Long, Cancel As ReturnBoolean, Control As Control, Action As fmAction, Data As DataObject, X As Single, Y As Single, Effect As ReturnEffect, Shift As Integer)`
- `Change()`
- `Click(Index As Long)`
- `DblClick(Index As Long, Cancel As ReturnBoolean)`
- `Error(Index As Long, Number As Integer, Description As ReturnString, SCode As Long, Source As String, HelpFile As String, HelpContext As Long, CancelDisplay As ReturnBoolean)`
- `KeyDown(KeyCode As ReturnInteger, Shift As Integer)`
- `KeyPress(KeyAscii As ReturnInteger)`
- `KeyUp(KeyCode As ReturnInteger, Shift As Integer)`
- `Layout(Index As Long)`
- `MouseDown(Index As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)`
- `MouseMove(Index As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)`
- `MouseUp(Index As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)`
- `RemoveControl(Index As Long, Control As Control)`
- `Scroll(Index As Long, ActionX As fmScrollAction, ActionY As fmScrollAction, RequestDx As Single, RequestDy As Single, ActualDx As ReturnSingle, ActualDy As ReturnSingle)`
- `Zoom(Index As Long, Percent As Integer)`
