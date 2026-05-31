# OptionFrameEvents

**Type:** Dispatch Interface  
**Library:** Microsoft Forms 2.0 Object Library  
**GUID:** {CF3F94A0-F546-11CE-9BCE-00AA00608E01}  

## Methods (16)

- `AddControl(Control As Control)`
- `BeforeDragOver(Cancel As ReturnBoolean, Control As Control, Data As DataObject, X As Single, Y As Single, State As fmDragState, Effect As ReturnEffect, Shift As Integer)`
- `BeforeDropOrPaste(Cancel As ReturnBoolean, Control As Control, Action As fmAction, Data As DataObject, X As Single, Y As Single, Effect As ReturnEffect, Shift As Integer)`
- `Click()`
- `DblClick(Cancel As ReturnBoolean)`
- `Error(Number As Integer, Description As ReturnString, SCode As Long, Source As String, HelpFile As String, HelpContext As Long, CancelDisplay As ReturnBoolean)`
- `KeyDown(KeyCode As ReturnInteger, Shift As Integer)`
- `KeyPress(KeyAscii As ReturnInteger)`
- `KeyUp(KeyCode As ReturnInteger, Shift As Integer)`
- `Layout()`
- `MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)`
- `MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)`
- `MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)`
- `RemoveControl(Control As Control)`
- `Scroll(ActionX As fmScrollAction, ActionY As fmScrollAction, RequestDx As Single, RequestDy As Single, ActualDx As ReturnSingle, ActualDy As ReturnSingle)`
- `Zoom(Percent As Integer)`
