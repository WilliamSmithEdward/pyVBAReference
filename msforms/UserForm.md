# UserForm

**Type:** Class  
**Library:** Microsoft Forms 2.0 Object Library  
**GUID:** {C62A69F0-16DC-11CE-9E98-00AA00574A4F}  

## Properties (31)

- `ActiveControl As Control  (read-only)`
- `BackColor As OLE_COLOR  (read/write)`
- `BorderColor As OLE_COLOR  (read/write)`
- `BorderStyle As fmBorderStyle  (read/write)`
- `CanPaste As Boolean  (read-only)`
- `CanRedo As Boolean  (read-only)`
- `CanUndo As Boolean  (read-only)`
- `Caption As String  (read/write)`
- `Controls As Controls  (read-only)`
- `Cycle As fmCycle  (read/write)`
- `Enabled As Boolean  (read/write)`
- `Font As Font  (read/write)`
- `ForeColor As OLE_COLOR  (read/write)`
- `InsideHeight As Single  (read-only)`
- `InsideWidth As Single  (read-only)`
- `KeepScrollBarsVisible As fmScrollBars  (read/write)`
- `MouseIcon As Picture  (read/write)`
- `MousePointer As fmMousePointer  (read/write)`
- `PictureAlignment As fmPictureAlignment  (read/write)`
- `Picture As Picture  (read/write)`
- `PictureSizeMode As fmPictureSizeMode  (read/write)`
- `PictureTiling As Boolean  (read/write)`
- `ScrollBars As fmScrollBars  (read/write)`
- `ScrollHeight As Single  (read/write)`
- `ScrollLeft As Single  (read/write)`
- `ScrollTop As Single  (read/write)`
- `ScrollWidth As Single  (read/write)`
- `SpecialEffect As fmSpecialEffect  (read/write)`
- `VerticalScrollBarSide As fmVerticalScrollBarSide  (read/write)`
- `Zoom As Integer  (read/write)`
- `DrawBuffer As Long  (read/write)`

## Methods (22)

- `_GetInsideHeight(InsideHeight As Long)`
- `_GetInsideWidth(InsideWidth As Long)`
- `_SetScrollHeight(ScrollHeight As Long)`
- `_GetScrollHeight(ScrollHeight As Long)`
- `_SetScrollLeft(ScrollLeft As Long)`
- `_GetScrollLeft(ScrollLeft As Long)`
- `_SetScrollTop(ScrollTop As Long)`
- `_GetScrollTop(ScrollTop As Long)`
- `_SetScrollWidth(ScrollWidth As Long)`
- `_GetScrollWidth(ScrollWidth As Long)`
- `Copy()`
- `Cut()`
- `Paste()`
- `RedoAction()`
- `Repaint()`
- `Scroll([xAction As Variant], [yAction As Variant])`
- `SetDefaultTabOrder()`
- `UndoAction()`
- `_SetGridX(GridX As Long)`
- `_GetGridX(GridX As Long)`
- `_SetGridY(GridY As Long)`
- `_GetGridY(GridY As Long)`

## Events (16)

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
