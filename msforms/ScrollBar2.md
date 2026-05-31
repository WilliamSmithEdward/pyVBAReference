# ScrollBar2

**Type:** Class  
**Library:** Microsoft Forms 2.0 Object Library  
**GUID:** {6240EF28-7EAB-4DC7-A5E3-7CFB35EFB34D}  

## Properties (13)

- `BackColor As OLE_COLOR  (read/write)`
- `ForeColor As OLE_COLOR  (read/write)`
- `Enabled As Boolean  (read/write)`
- `MousePointer As fmMousePointer  (read/write)`
- `Value As Long  (read/write)`
- `Min As Long  (read/write)`
- `Max As Long  (read/write)`
- `SmallChange As Long  (read/write)`
- `LargeChange As Long  (read/write)`
- `ProportionalThumb As Boolean  (read/write)`
- `Orientation As fmOrientation  (read/write)`
- `Delay As Long  (read/write)`
- `MouseIcon As Picture  (read/write)`

## Events (8)

- `BeforeDragOver(Cancel As ReturnBoolean, Data As DataObject, X As Single, Y As Single, DragState As fmDragState, Effect As ReturnEffect, Shift As Integer)`
- `BeforeDropOrPaste(Cancel As ReturnBoolean, Action As fmAction, Data As DataObject, X As Single, Y As Single, Effect As ReturnEffect, Shift As Integer)`
- `Change()`
- `Error(Number As Integer, Description As ReturnString, SCode As Long, Source As String, HelpFile As String, HelpContext As Long, CancelDisplay As ReturnBoolean)`
- `KeyDown(KeyCode As ReturnInteger, Shift As Integer)`
- `KeyPress(KeyAscii As ReturnInteger)`
- `KeyUp(KeyCode As ReturnInteger, Shift As Integer)`
- `Scroll()`
