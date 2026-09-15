# WebBrowserControl

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {D303AC37-74DB-45B9-8C22-AD7C3FBA68EF}  

Represents a web browser control on a form.

**Remarks:** Use the ControlSource property to specify the webpage to load in the control.

## Properties (72)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `OldValue As Variant  (read-only)`  
  Use the OldValue property to determine the unedited value of a bound control. Read-only Variant.
- `Properties As Properties  (read-only)`  
  Returns a reference to a control's Properties collection object. Read-only.
- `Controls As Children  (read-only)`  
  Returns the Controls collection of a form, subform, report, or section. Read-only Controls.
- `Hyperlink As _Hyperlink  (read-only)`  
  Use the Hyperlink property to return a reference to a Hyperlink object and to access the properties and methods of a control's hyperlink. Read-only.
- `Value As Variant  (read/write)`  
  Determines or specifies the text in the control. Read/write Variant.
- `ControlSource As String  (read/write)`  
  Use the ControlSource property to specify what data appears in a control. You can display and edit data bound to a field in a table, query, or SQL statement. You can also display the result of an expression. Read/write String.
- `DisplayWhen As Byte  (read/write)`  
  Use the DisplayWhen property to specify which of a form's controls you want displayed on screen and in print. Read/write Byte.
- `Enabled As Boolean  (read/write)`  
  Use the Enabled property to set or return the status of the conditional format in the FormatCondition object. Read/write Boolean.
- `ReadyState As AcWebBrowserState  (read-only)`  
  Gets the status of the specified web browser control. Read-only AcWebBrowserState.
- `Progress As Long  (read-only)`  
  Specifies the amount of total progress of a download operation. Read-only Long.
- `ScrollBars As AcWebBrowserScrollBars  (read/write)`  
  Use the ScrollBars property to specify whether scroll bars appear on a control. Read/write Byte.
- `ScrollTop As Long  (read/write)`  
  Gets or sets the distance, in pixels, between the top edge of the WebBrowserControl object and the topmost portion of the content currently visible in the control. Read/write Long.
- `ScrollLeft As Long  (read/write)`  
  Gets or sets the distance, in pixels, between the left edge of the WebBrowserControl object and the leftmost portion of the content currently visible in the control. Read/write Long.
- `Transform As String  (read/write)`  
  Read/write.
- `LocationURL As String  (read-only)`  
  Gets the Uniform Resource Locator (URL) of the current document. Read-only String.
- `Object As Object  (read-only)`  
  Use the Object property in Visual Basic to return a reference to the ActiveX object that is associated with a linked or embedded OLE object in a control. By using this reference, you can access the properties or invoke the methods of the OLE object. Read-only Object.
- `Left As Integer  (read/write)`  
  Use the Left property to specify an object's location on a form or report. Read/write Integer.
- `Top As Integer  (read/write)`  
  Use the Top property to specify an object's location on a form or report. Read/write Integer.
- `Width As Integer  (read/write)`  
  Gets or sets the width of the specified object in twips. Read/write Integer.
- `Height As Integer  (read/write)`  
  Gets or sets the height of the specified object in twips. Read/write Integer.
- `SpecialEffect As Integer  (read/write)`  
  Use the SpecialEffect property to specify whether special formatting will apply to the specified object. Read/write Byte.
- `BorderStyle As Byte  (read/write)`  
  Specifies how a control's border appears. Read/write Byte.
- `BorderColor As Long  (read/write)`  
  Use the BorderColor property to specify the color of a control's border. Read/write Long.
- `BorderWidth As Byte  (read/write)`  
  Use the BorderWidth property to specify the width of a control's border. Read/write Byte.
- `Section As Integer  (read/write)`  
  You can identify these controls by the section of a form or report where the control appears. Read/write Integer.
- `OnMouseDown As String  (read/write)`  
  Sets or returns the value of the On Mouse Down box in the Properties window. Read/write String.
- `OnMouseMove As String  (read/write)`  
  Sets or returns the value of the On Mouse Move box in the Properties window. Read/write String.
- `OnMouseUp As String  (read/write)`  
  Sets or returns the value of the On Mouse Up box in the Properties window. Read/write String.
- `OnKeyDown As String  (read/write)`  
  Sets or returns the value of the On Key Down box in the Properties window. Read/write String.
- `OnKeyUp As String  (read/write)`  
  Sets or returns the value of the On Key Up box in the Properties window. Read/write String.
- `OnKeyPress As String  (read/write)`  
  Sets or returns the value of the On Key Press box in the Properties window. Read/write String.
- `Name As String  (read/write)`  
  Use the Name property to specify or determine the string expression that identifies the name of an object. Read/write String.
- `Layout As AcLayoutType  (read-only)`  
  Returns the type of layout for the specified control. Read-only AcLayoutType.
- `LeftPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the control and its left gridline. Read/write Integer.
- `TopPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the control and its top gridline. Read/write Integer.
- `RightPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the control and its right gridline. Read/write Integer.
- `BottomPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the control and its bottom gridline. Read/write Integer.
- `GridlineStyleLeft As Byte  (read/write)`  
  Gets or sets the left gridline style of the specified control. Read/write Byte.
- `GridlineStyleTop As Byte  (read/write)`  
  Gets or sets the top gridline style of the specified control. Read/write Byte.
- `GridlineStyleRight As Byte  (read/write)`  
  Gets or sets the right gridline style of the specified control. Read/write Byte.
- `GridlineStyleBottom As Byte  (read/write)`  
  Gets or sets the bottom gridline style of the specified control. Read/write Byte.
- `GridlineWidthLeft As Byte  (read/write)`  
  Gets or sets the width of the left gridline for the specified control. Read/write Byte.
- `GridlineWidthTop As Byte  (read/write)`  
  Gets or sets the width of the top gridline for the specified control. Read/write Byte.
- `GridlineWidthRight As Byte  (read/write)`  
  Gets or sets the width of the right gridline for the specified control. Read/write Byte.
- `GridlineWidthBottom As Byte  (read/write)`  
  Gets or sets the width of the bottom gridline for the specified control. Read/write Byte.
- `GridlineColor As Long  (read/write)`  
  Gets or sets the color of the gridline for the specified control. Read/write Long.
- `HorizontalAnchor As AcHorizontalAnchor  (read/write)`  
  Gets or sets an AcHorizontalAnchor constant that indicates how the control is anchored horizontally within its layout. Read/write.
- `VerticalAnchor As AcVerticalAnchor  (read/write)`  
  Gets or sets an AcVerticalAnchor constant that indicates how the specified control is anchored vertically within its layout. Read/write.
- `LayoutID As Long  (read-only)`  
  Returns the unique identifier for the layout that contains the specified control. Read-only Long.
- `OnBeforeNavigate As String  (read/write)`  
  Gets or sets the value of the On Before Navigate box in the property sheet of a web browser control. Read/write String.
- `OnDocumentComplete As String  (read/write)`  
  Gets or sets the value of the On Document Complete box in the property sheet of a web browser control. Read/write String.
- `OnNavigateError As String  (read/write)`  
  Gets or sets the value of the On Navigate Error box in the property sheet of a web browser control. Read/write String.
- `OnProgressChange As String  (read/write)`  
  Gets or sets the value of the On Progress Change box in the property sheet of a web browser control. Read/write String.
- `OnUpdated As String  (read/write)`  
  Sets or returns the value of the On Updated box in the Properties window of a form or report. Read/write String.
- `TabStop As Boolean  (read/write)`  
  Use the TabStop property to specify whether you can use the Tab key to move the focus to a control. Read/write Boolean.
- `TabIndex As Integer  (read/write)`  
  Use the TabIndex property to specify a control's place in the tab order on a form or report. Read/write Integer.
- `Visible As Boolean  (read/write)`  
  Returns or sets whether the object is visible. Read/write Boolean.
- `InSelection As Boolean  (read/write)`  
  Use the InSelection property to determine or specify whether a control on a form in Design view is selected. Read/write Boolean.
- `ControlType As Byte  (read/write)`  
  Use the ControlType property in Visual Basic to determine the type of control on a form or report. Read/write Byte.
- `BorderThemeColorIndex As Long  (read/write)`  
  Gets or sets a value that represents a color in the applied color theme associated with the BorderColor property of the specified object. Read/write Long.
- `BorderTint As Single  (read/write)`  
  Gets or sets the tint that is applied to the theme color in the BorderColor property of the specified object. Read/write Single.
- `BorderShade As Single  (read/write)`  
  Gets or sets the shade applied to the theme color in the BorderColor property of the specified object. Read/write Single.
- `GridlineThemeColorIndex As Long  (read/write)`  
  Gets or sets the theme color index that represents a color in the applied color theme associated with the GridlineColor property of the specified object. Read/write Long.
- `GridlineTint As Single  (read/write)`  
  Gets or sets the tint applied to the theme color in the GridlineColor property of the specified object. Read/write Single.
- `GridlineShade As Single  (read/write)`  
  Gets or sets the shade applied to the theme color in the GridlineColor property of the specified object. Read/write Single.
- `StatusBarText As String  (read/write)`  
  Use the StatusBarText property to specify the text that is displayed in the status bar when a control is selected. Read/write String.
- `ControlTipText As String  (read/write)`  
  Use the ControlTipText property to specify the text that appears in a ScreenTip when you hold the mouse pointer over a control. Read/write String.
- `EventProcPrefix As String  (read/write)`  
  Gets or sets the prefix portion of an event procedure name. Read/write String.
- `HelpContextId As Long  (read/write)`  
  The HelpContextID property specifies the context ID of a topic in the custom Help file specified by the HelpFile property setting. Read/write Long.
- `Tag As String  (read/write)`  
  Stores extra information about a form, report, section, or control needed by a Microsoft Access application. Read/write String.

## Methods (5)

- `Undo()`  
  Use the Undo method to reset a control or form when its value has been changed.
- `SizeToFit()`  
  Use the SizeToFit method to size a control so that it fits the text or image that it contains.
- `Requery()`  
  The Requery method updates the data underlying a specified control that's on the active form by requerying the source of data for the control.
- `SetFocus()`  
  The SetFocus method moves the focus to the specified form, the specified control on the active form, or the specified field on the active datasheet.
- `Move(Left As Variant, [Top As Variant], [Width As Variant], [Height As Variant])`  
  Moves the specified object to the coordinates specified by the argument values.
    - `Left As Variant` (required): The screen position in twips for the left edge of the object relative to the left edge of the Microsoft Access window.
    - `Top As Variant` (optional): The screen position in twips for the top edge of the object relative to the top edge of the Access window.
    - `Width As Variant` (optional): The desired width of the object in twips.
    - `Height As Variant` (optional): The desired height of the object in twips.

## Events (21)

- `Updated(Code As Integer)`  
  The Updated event occurs when an OLE object's data has been modified.
- `BeforeUpdate(Cancel As Integer)`  
  The BeforeUpdate event occurs before changed data in a control or record is updated.
    - `Cancel As Integer` (required): The setting determines if the BeforeUpdate event occurs. Setting the Cancel argument to True (1) cancels the BeforeUpdate event.
- `AfterUpdate()`  
  The AfterUpdate event occurs after changed data in a control or record is updated.
- `Enter()`  
  The Enter event occurs before a control actually receives the focus from a control on the same form or report.
- `Exit(Cancel As Integer)`  
  The Exit event occurs just before a control loses the focus to another control on the same form or report.
    - `Cancel As Integer` (required): Set to True to cancel the event.
- `Dirty(Cancel As Integer)`  
  The Dirty event occurs when the contents of the specified control changes.
    - `Cancel As Integer` (required): The setting determines if the Dirty event occurs. Setting the _Cancel_ argument to True (1) cancels the Dirty event.
- `Change()`  
  The Change event occurs when the contents of the specified control change.
- `GotFocus()`  
  The GotFocus event occurs when the specified object receives the focus.
- `LostFocus()`  
  The LostFocus event occurs when the specified object loses the focus.
- `Click()`  
  The Click event occurs when the user presses and then releases a mouse button over an object.
- `DblClick(Cancel As Integer)`  
  The DblClick event occurs when the user presses and releases the left mouse button twice over an object within the double-click time limit of the system.
    - `Cancel As Integer` (required): The setting determines if the DblClick event occurs. Setting the _Cancel_ argument to True (1) cancels the DblClick event.
- `MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)`  
  The MouseDown event occurs when the user presses a mouse button.
    - `Button As Integer` (required): The button that was pressed to trigger the event. If you need to test for the _Button_ argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acLeftButton</b> The bit mask for the left mouse button.</p></li><li><p><b>acRightButton</b> The bit mask for the right mouse button.</p></li><li><p><b>acMiddleButton</b> The bit mask for the middle mouse button.</p></li></ul>
    - `Shift As Integer` (required): The state of the Shift, Ctrl, and Alt keys when the button specified by the _Button_ argument was pressed or released. If you need to test for the _Shift_ argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acShiftMask</b> The bit mask for the Shift key.</p></li><li><p><b>acCtrlMask</b> The bit mask for the Ctrl key.</p></li><li><p><b>acAltMask</b> The bit mask for the Alt key.</p></li></ul>
    - `X As Single` (required): The _x_ coordinate for the current location of the mouse pointer, in twips.
    - `Y As Single` (required): The _y_ coordinate for the current location of the mouse pointer, in twips.
- `MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)`  
  The MouseMove event occurs when the user moves the mouse.
    - `Button As Integer` (required): The button that was pressed or released when the event was triggered. If you need to test for the _Button_ argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acLeftButton</b> The bit mask for the left mouse button.</p></li><li><p><b>acRightButton</b> The bit mask for the right mouse button.</p></li><li><p><b>acMiddleButton</b> The bit mask for the middle mouse button.</p></li></ul>
    - `Shift As Integer` (required): The state of the Shift, Ctrl, and Alt keys when the button specified by the _Button_ argument was pressed or released. If you need to test for the _Shift_ argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acShiftMask</b> The bit mask for the Shift key.</p></li><li><p><b>acCtrlMask</b> The bit mask for the Ctrl key.</p></li><li><p><b>acAltMask</b> The bit mask for the Alt key.</p></li></ul>
    - `X As Single` (required): The _x_ coordinate for the current location of the mouse pointer, in twips.
    - `Y As Single` (required): The _y_ coordinate for the current location of the mouse pointer, in twips.
- `MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)`  
  The MouseUp event occurs when the user releases a mouse button.
    - `Button As Integer` (required): The button that was released to trigger the event. If you need to test for the _Button_ argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acLeftButton</b> The bit mask for the left mouse button.</p></li><li><p><b>acRightButton</b> The bit mask for the right mouse button.</p></li><li><p><b>acMiddleButton</b> The bit mask for the middle mouse button.</p></li></ul>
    - `Shift As Integer` (required): The state of the Shift, Ctrl, and Alt keys when the button specified by the _Button_ argument was pressed or released. If you need to test for the _Shift_ argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acShiftMask</b> The bit mask for the Shift key.</p></li><li><p><b>acCtrlMask</b> The bit mask for the Ctrl key.</p></li><li><p><b>acAltMask</b> The bit mask for the Alt key.</p></li></ul>
    - `X As Single` (required): The _x_ coordinate for the current location of the mouse pointer, in twips.
    - `Y As Single` (required): The _y_ coordinate for the current location of the mouse pointer, in twips.
- `KeyDown(KeyCode As Integer, Shift As Integer)`  
  The KeyDown event occurs when the user presses a key while a form or control has the focus. This event also occurs if you send a keystroke to a form or control by using the SendKeys action in a macro or the SendKeys statement in Visual Basic.
    - `KeyCode As Integer` (required): A key code, such as vbKeyF1 (the F1 key) or vbKeyHome (the Home key). To specify key codes, use the intrinsic constants shown in the Object Browser. You can prevent an object from receiving a keystroke by setting _KeyCode_ to 0.
    - `Shift As Integer` (required): The state of the Shift, Ctrl, and Alt keys at the time of the event. If you need to test for the _Shift_ argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p>acShiftMask The bit mask for the Shift key.</p></li><li><p>acCtrlMask The bit mask for the Ctrl key.</p></li><li><p>acAltMask The bit mask for the Alt key.</p></li></ul>
- `KeyPress(KeyAscii As Integer)`  
  The KeyPress event occurs when the user presses and releases a key or key combination that corresponds to an ANSI code while a form or control has the focus. This event also occurs if you send an ANSI keystroke to a form or control by using the SendKeys action in a macro or the SendKeys statement in Visual Basic.
    - `KeyAscii As Integer` (required): Returns a numeric ANSI key code. The _KeyAscii_ argument is passed by reference; changing it sends a different character to the object. Setting the _KeyAscii_ argument to 0 cancels the keystroke so that the object doesn't recognize that a key was pressed.
- `KeyUp(KeyCode As Integer, Shift As Integer)`  
  The KeyUp event occurs when the user releases a key while a form or control has the focus. This event also occurs if you send a keystroke to a form or control by using the SendKeys action in a macro or the SendKeys statement in Visual Basic.
    - `KeyCode As Integer` (required): A key code, such as vbKeyF1 (the F1 key) or vbKeyHome (the Home key). To specify key codes, use the intrinsic constants shown in the Object Browser. You can prevent an object from receiving a keystroke by setting _KeyCode_ to 0.
    - `Shift As Integer` (required): The state of the Shift, Ctrl, and Alt keys at the time of the event. If you need to test for the _Shift_ argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acShiftMask</b> The bit mask for the Shift key.</p></li><li><p><b>acCtrlMask</b> The bit mask for the Ctrl key.</p></li><li><p><b>acAltMask</b> The bit mask for the Alt key.</p></li></ul>
- `BeforeNavigate2(pDisp As Object, URL As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)`  
  Occurs before navigation occurs in the given WebBrowserControl.
    - `pDisp As Object` (required): A pointer to the IDispatch interface for the WebBrowserControl object that represents the window or frame.
    - `URL As Variant` (required): Contains the URL to be navigated to.
    - `flags As Variant` (required): Reserved. Must be set to Null.
    - `TargetFrameName As Variant` (required): Contains the name of the frame in which to display the resource, or Null if no named frame is targeted for the resource.
    - `PostData As Variant` (required): Contains the data to send to the server, if the HTTP POST transaction is used.
    - `Headers As Variant` (required): Contains additional HTTP headers to send to the server (HTTP URLs only). The headers can specify information, such as the action required of the server, the type of data being passed to the server, or a status code.
    - `Cancel As Boolean` (required): Contains the cancel flag. Set to True to cancel the navigation operation.
- `DocumentComplete(pDisp As Object, URL As Variant)`  
  Occurs when a document is completely loaded and initialized.
    - `pDisp As Object` (required): A pointer to the IDispatch interface of the window or frame in which the document is loaded.
    - `URL As Variant` (required): Contains the URL of the loaded document.
- `ProgressChange(Progress As Long, ProgressMax As Long)`  
  Occurs when the progress of a download operation is updated.
    - `Progress As Long` (required): Specifies the amount of total progress to show, or -1 when progress is complete.
    - `ProgressMax As Long` (required): Specifies the maximum progress value.
- `NavigateError(pDisp As Object, URL As Variant, TargetFrameName As Variant, StatusCode As Variant, Cancel As Boolean)`  
  Occurs when an error occurs during navigation.
    - `pDisp As Object` (required): A pointer to an IDispatch interface for the WebBrowserControl object that represents the window or frame in which the navigation error occurred.
    - `URL As Variant` (required): Contains the URL for which navigation failed.
    - `TargetFrameName As Variant` (required): Contains the name of the frame in which to display the resource, or Null if no named frame was targeted for the resource.
    - `StatusCode As Variant` (required): Contains an error status code, if available.
    - `Cancel As Boolean` (required): Specifies whether to cancel the navigation to an error page or to any further autosearch.
