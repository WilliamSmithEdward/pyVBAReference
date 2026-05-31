# TabStrip

**Type:** Class  
**Library:** Microsoft Forms 2.0 Object Library  
**GUID:** {EAE50EB0-4A62-11CE-BED6-00AA00611080}  

## Properties (18)

- `BackColor As OLE_COLOR  (read/write)`
- `ForeColor As OLE_COLOR  (read/write)`
- `Font As Font  (read/write)`
- `Enabled As Boolean  (read/write)`
- `MouseIcon As Picture  (read/write)`
- `MousePointer As fmMousePointer  (read/write)`
- `MultiRow As Boolean  (read/write)`
- `Style As fmTabStyle  (read/write)`
- `TabOrientation As fmTabOrientation  (read/write)`
- `ClientTop As Single  (read-only)`
- `ClientLeft As Single  (read-only)`
- `ClientWidth As Single  (read-only)`
- `ClientHeight As Single  (read-only)`
- `Tabs As Tabs  (read-only)`
- `SelectedItem As Tab  (read-only)`
- `Value As Long  (read/write)`
- `TabFixedWidth As Single  (read/write)`
- `TabFixedHeight As Single  (read/write)`

## Methods (8)

- `_SetTabFixedWidth(TabFixedWidth As Long)`
- `_GetTabFixedWidth(TabFixedWidth As Long)`
- `_SetTabFixedHeight(TabFixedHeight As Long)`
- `_GetTabFixedHeight(TabFixedHeight As Long)`
- `_GetClientTop(ClientTop As Long)`
- `_GetClientLeft(ClientLeft As Long)`
- `_GetClientWidth(ClientWidth As Long)`
- `_GetClientHeight(ClientHeight As Long)`

## Events (12)

- `BeforeDragOver(Index As Long, Cancel As ReturnBoolean, Data As DataObject, X As Single, Y As Single, DragState As fmDragState, Effect As ReturnEffect, Shift As Integer)`
- `BeforeDropOrPaste(Index As Long, Cancel As ReturnBoolean, Action As fmAction, Data As DataObject, X As Single, Y As Single, Effect As ReturnEffect, Shift As Integer)`
- `Change()`
- `Click(Index As Long)`
- `DblClick(Index As Long, Cancel As ReturnBoolean)`
- `Error(Number As Integer, Description As ReturnString, SCode As Long, Source As String, HelpFile As String, HelpContext As Long, CancelDisplay As ReturnBoolean)`
- `KeyDown(KeyCode As ReturnInteger, Shift As Integer)`
- `KeyPress(KeyAscii As ReturnInteger)`
- `KeyUp(KeyCode As ReturnInteger, Shift As Integer)`
- `MouseDown(Index As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)`
- `MouseMove(Index As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)`
- `MouseUp(Index As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)`
