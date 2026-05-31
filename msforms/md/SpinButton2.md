# SpinButton2

**Type:** Class  
**Library:** Microsoft Forms 2.0 Object Library  
**GUID:** {EA778DB4-CE69-4DA5-BC1D-34E2168D5EED}  

## Properties (11)

- `BackColor As OLE_COLOR  (read/write)`
- `ForeColor As OLE_COLOR  (read/write)`
- `Enabled As Boolean  (read/write)`
- `MousePointer As fmMousePointer  (read/write)`
- `Value As Long  (read/write)`
- `Min As Long  (read/write)`
- `Max As Long  (read/write)`
- `SmallChange As Long  (read/write)`
- `Orientation As fmOrientation  (read/write)`
- `MouseIcon As Picture  (read/write)`
- `Delay As Long  (read/write)`

## Events (9)

- `BeforeDragOver(Cancel As ReturnBoolean, Data As DataObject, X As Single, Y As Single, DragState As fmDragState, Effect As ReturnEffect, Shift As Integer)`
- `BeforeDropOrPaste(Cancel As ReturnBoolean, Action As fmAction, Data As DataObject, X As Single, Y As Single, Effect As ReturnEffect, Shift As Integer)`
- `Change()`
- `Error(Number As Integer, Description As ReturnString, SCode As Long, Source As String, HelpFile As String, HelpContext As Long, CancelDisplay As ReturnBoolean)`
- `KeyDown(KeyCode As ReturnInteger, Shift As Integer)`
- `KeyPress(KeyAscii As ReturnInteger)`
- `KeyUp(KeyCode As ReturnInteger, Shift As Integer)`
- `SpinUp()`
- `SpinDown()`
