# DispWebBrowserControlEvents

**Type:** Dispatch Interface  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {EACB9075-68F8-4E3B-B865-E1CE6BE0447C}  

## Methods (21)

- `Updated(Code As Integer)`
- `BeforeUpdate(Cancel As Integer)`
- `AfterUpdate()`
- `Enter()`
- `Exit(Cancel As Integer)`
- `Dirty(Cancel As Integer)`
- `Change()`
- `GotFocus()`
- `LostFocus()`
- `Click()`
- `DblClick(Cancel As Integer)`
- `MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)`
- `MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)`
- `MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)`
- `KeyDown(KeyCode As Integer, Shift As Integer)`
- `KeyPress(KeyAscii As Integer)`
- `KeyUp(KeyCode As Integer, Shift As Integer)`
- `BeforeNavigate2(pDisp As Object, URL As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)`
- `DocumentComplete(pDisp As Object, URL As Variant)`
- `ProgressChange(Progress As Long, ProgressMax As Long)`
- `NavigateError(pDisp As Object, URL As Variant, TargetFrameName As Variant, StatusCode As Variant, Cancel As Boolean)`
