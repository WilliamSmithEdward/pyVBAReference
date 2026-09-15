# Edge

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {3B06F000-E47C-11CD-8701-00AA003F0F07}  

## Properties (68)

- `Application As Application  (read-only)`
- `Parent As Object  (read-only)`
- `OldValue As Variant  (read-only)`
- `Properties As Properties  (read-only)`
- `Controls As Children  (read-only)`
- `Hyperlink As _Hyperlink  (read-only)`
- `Value As Variant  (read/write)`
- `ControlSource As String  (read/write)`
- `DisplayWhen As Byte  (read/write)`
- `Enabled As Boolean  (read/write)`
- `ReadyState As AcWebBrowserState  (read-only)`
- `ScrollBars As AcWebBrowserScrollBars  (read/write)`
- `ScrollTop As Long  (read/write)`
- `ScrollLeft As Long  (read/write)`
- `LocationURL As String  (read-only)`
- `Left As Integer  (read/write)`
- `Top As Integer  (read/write)`
- `Width As Integer  (read/write)`
- `Height As Integer  (read/write)`
- `SpecialEffect As Integer  (read/write)`
- `BorderStyle As Byte  (read/write)`
- `BorderColor As Long  (read/write)`
- `BorderWidth As Byte  (read/write)`
- `Section As Integer  (read/write)`
- `OnMouseDown As String  (read/write)`
- `OnMouseMove As String  (read/write)`
- `OnMouseUp As String  (read/write)`
- `OnKeyDown As String  (read/write)`
- `OnKeyUp As String  (read/write)`
- `OnKeyPress As String  (read/write)`
- `Name As String  (read/write)`
- `Layout As AcLayoutType  (read-only)`
- `LeftPadding As Integer  (read/write)`
- `TopPadding As Integer  (read/write)`
- `RightPadding As Integer  (read/write)`
- `BottomPadding As Integer  (read/write)`
- `GridlineStyleLeft As Byte  (read/write)`
- `GridlineStyleTop As Byte  (read/write)`
- `GridlineStyleRight As Byte  (read/write)`
- `GridlineStyleBottom As Byte  (read/write)`
- `GridlineWidthLeft As Byte  (read/write)`
- `GridlineWidthTop As Byte  (read/write)`
- `GridlineWidthRight As Byte  (read/write)`
- `GridlineWidthBottom As Byte  (read/write)`
- `GridlineColor As Long  (read/write)`
- `HorizontalAnchor As AcHorizontalAnchor  (read/write)`
- `VerticalAnchor As AcVerticalAnchor  (read/write)`
- `LayoutID As Long  (read-only)`
- `OnBeforeNavigate As String  (read/write)`
- `OnDocumentComplete As String  (read/write)`
- `OnNavigateError As String  (read/write)`
- `OnUpdated As String  (read/write)`
- `TabStop As Boolean  (read/write)`
- `TabIndex As Integer  (read/write)`
- `Visible As Boolean  (read/write)`
- `InSelection As Boolean  (read/write)`
- `ControlType As Byte  (read/write)`
- `BorderThemeColorIndex As Long  (read/write)`
- `BorderTint As Single  (read/write)`
- `BorderShade As Single  (read/write)`
- `GridlineThemeColorIndex As Long  (read/write)`
- `GridlineTint As Single  (read/write)`
- `GridlineShade As Single  (read/write)`
- `StatusBarText As String  (read/write)`
- `ControlTipText As String  (read/write)`
- `EventProcPrefix As String  (read/write)`
- `HelpContextId As Long  (read/write)`
- `Tag As String  (read/write)`

## Methods (8)

- `SizeToFit()`
- `Requery()`
- `SetFocus()`
- `Move(Left As Variant, [Top As Variant], [Width As Variant], [Height As Variant])`
- `ExecuteJavascript(Script As String)`
- `RetrieveJavascriptValue(Expression As String) As String`
- `Navigate(URL As String)`
- `Refresh()`

## Events (20)

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
- `BeforeNavigate(Cancel As Integer, URL As String)`
- `DocumentComplete(URL As Variant)`
- `NavigateError(URL As String, StatusCode As Variant)`
