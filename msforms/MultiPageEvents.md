# MultiPageEvents

**Type:** Dispatch Interface  
**Library:** Microsoft Forms 2.0 Object Library  
**GUID:** {7B020EC8-AF6C-11CE-9F46-00AA00574A4F}  

## Methods (17)

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
