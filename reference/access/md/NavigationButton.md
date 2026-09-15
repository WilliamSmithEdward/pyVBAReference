# NavigationButton

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {3B06E993-E47C-11CD-8701-00AA003F0F07}  

This object represents a navigation button in a navigation control on a form.

**Remarks:** When you choose a navigation button, the form or report specified in the NavigationTargetName is displayed. Use the NavigationWhereClause to filter the records displayed on the form. Use the Tabs property of the NavigationControl object to return the collection of navigation buttons for a navigation control. Use the SelectedTab property of the NavigationControl object to return the navigation button that is currently selected.

## Properties (126)

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
- `EventProcPrefix As String  (read/write)`  
  Gets or sets the prefix portion of an event procedure name. Read/write String.
- `ControlType As Byte  (read/write)`  
  Use the ControlType property in Visual Basic to determine the type of control on a form or report. Read/write Byte.
- `Caption As String  (read/write)`  
  Gets or sets the text that appears in the control. Read/write String.
- `Picture As String  (read/write)`  
  Use the Picture property to specify a bitmap or other type of graphic to be displayed on the specified control. Read/write String.
- `PictureType As Byte  (read/write)`  
  Use the PictureType property to specify whether Microsoft Access stores an object's picture as a linked or an embedded object. Read/write Byte.
- `PictureData As Variant  (read/write)`  
  Use the PictureData property to copy the picture to another object that supports the Picture property. Read/write Variant.
- `Transparent As Boolean  (read/write)`  
  Use the Transparent property to specify whether a command button is solid or transparent. Read/write Boolean.
- `AutoRepeat As Boolean  (read/write)`  
  Use the AutoRepeat property to specify whether an event procedure or macro runs repeatedly while a command button on a form remains pressed in. Read/write Boolean.
- `StatusBarText As String  (read/write)`  
  Use the StatusBarText property to specify the text that is displayed in the status bar when a control is selected. Read/write String.
- `OnPush As String  (read/write)`  
  Sets or returns the value of the On Click box in the Properties window of a command button. Read/write String.
- `HyperlinkAddress As String  (read/write)`  
  Use the HyperlinkAddress property to specify or determine the path to an object, document, webpage, or other destination for a hyperlink associated with the specified control. Read/write String.
- `HyperlinkSubAddress As String  (read/write)`  
  Use the HyperlinkSubAddress property to specify or determine a location within the target document specified by the HyperlinkAddress property. Read/write String.
- `Visible As Boolean  (read/write)`  
  Returns or sets whether the object is visible. Read/write Boolean.
- `DisplayWhen As Byte  (read/write)`  
  Use the DisplayWhen property to specify which of a form's controls you want displayed on screen and in print. Read/write Byte.
- `Enabled As Boolean  (read/write)`  
  Use the Enabled property to set or return the status of the conditional format in the FormatCondition object. Read/write Boolean.
- `TabStop As Boolean  (read/write)`  
  Use the TabStop property to specify whether you can use the Tab key to move the focus to a control. Read/write Boolean.
- `TabIndex As Integer  (read/write)`  
  Use the TabIndex property to specify a control's place in the tab order on a form or report. Read/write Integer.
- `Left As Integer  (read-only)`  
  Use the Left property to specify an object's location on a form or report. Read/write Integer.
- `Top As Integer  (read-only)`  
  Use the Top property to specify an object's location on a form or report. Read/write Integer.
- `Width As Integer  (read/write)`  
  Gets or sets the width of the specified object in twips. Read/write Integer.
- `Height As Integer  (read/write)`  
  Gets or sets the height of the specified object in twips. Read/write Integer.
- `ForeColor As Long  (read/write)`  
  Use the ForeColor property to specify the color for text in a control. Read/write Long.
- `FontName As String  (read/write)`  
  Use the FontName property to specify the font for text in the following situations:
- `FontSize As Integer  (read/write)`  
  Use the FontSize property to specify the point size for text in the following situations:
- `FontWeight As Integer  (read/write)`  
  Use the FontWeight property to specify the line width that Windows uses to display and print characters in a control. Read/write Integer.
- `FontItalic As Boolean  (read/write)`  
  Use the FontItalic property to specify whether text is italic in the following situations:
- `FontUnderline As Boolean  (read/write)`  
  Use the FontUnderline property to specify whether text is underlined in the following situations:
- `FontBold As Integer  (read/write)`  
  Use the FontBold property to specify whether a font appears in a bold style in the following situations:
- `ShortcutMenuBar As String  (read/write)`  
  Use the ShortcutMenuBar property to specify the shortcut menu that appears when you right-click the specified object. Read/write String.
- `ControlTipText As String  (read/write)`  
  Use the ControlTipText property to specify the text that appears in a ScreenTip when you hold the mouse pointer over a control. Read/write String.
- `HelpContextId As Long  (read/write)`  
  The HelpContextID property specifies the context ID of a topic in the custom Help file specified by the HelpFile property setting. Read/write Long.
- `AutoLabel As Boolean  (read/write)`  
  Specifies whether labels are automatically created and attached to new controls. Read/write Boolean.
- `AddColon As Boolean  (read/write)`  
  Specifies whether a colon follows the text in labels for new controls. Read/write Boolean.
- `LabelX As Integer  (read/write)`  
  The LabelX property (along with the LabelY property) specifies the placement of the label for a new control. Read/write Integer.
- `LabelY As Integer  (read/write)`  
  The LabelY property (along with the LabelX property) specifies the placement of the label for a new control. Read/write Integer.
- `LabelAlign As Byte  (read/write)`  
  The LabelAlign property specifies the text alignment within attached labels on new controls. Read/write Byte.
- `Section As Integer  (read/write)`  
  You can identify these controls by the section of a form or report where the control appears. Read/write Integer.
- `Tag As String  (read/write)`  
  Stores extra information about a form, report, section, or control needed by a Microsoft Access application. Read/write String.
- `ObjectPalette As Variant  (read/write)`  
  The ObjectPalette property specifies the palette in the application used to create a bitmap or other graphic that is loaded into the specified control by using the Picture property. Read/write Variant.
- `IsVisible As Boolean  (read/write)`  
  Use the IsVisible property to determine whether a control on a report is visible. Read/write Boolean.
- `InSelection As Boolean  (read/write)`  
  Use the InSelection property to determine or specify whether a control on a form in Design view is selected. Read/write Boolean.
- `OnEnter As String  (read/write)`  
  Sets or returns the value of the On Enter box in the Properties window of specified object. Read/write String.
- `OnExit As String  (read/write)`  
  Sets or returns the value of the On Exit box in the Properties window of specified object. Read/write String.
- `OnGotFocus As String  (read/write)`  
  Sets or returns the value of the On Got Focus box in the Properties window of the specified object. Read/write String.
- `OnLostFocus As String  (read/write)`  
  Sets or returns the value of the On Lost Focus box in the Properties window of the specified object. Read/write String.
- `OnClick As String  (read/write)`  
  Sets or returns the value of the On Click box in the Properties window. Read/write String.
- `OnDblClick As String  (read/write)`  
  Sets or returns the value of the On Dbl Click box in the Properties window. Read/write String.
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
- `ReadingOrder As Byte  (read/write)`  
  Use the ReadingOrder property to specify or determine the reading order of words in text. Read/write Byte.
- `Name As String  (read/write)`  
  Use the Name property to specify or determine the string expression that identifies the name of an object. Read/write String.
- `Layout As AcLayoutType  (read-only)`  
  Returns the type of layout for the specified button. Read-only AcLayoutType.
- `LeftPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the button and its left gridline. Read/write Integer.
- `TopPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the button and its top gridline. Read/write Integer.
- `RightPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the button and its right gridline. Read/write Integer.
- `BottomPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the button and its bottom gridline. Read/write Integer.
- `GridlineStyleLeft As Byte  (read/write)`  
  Gets or sets the left gridline style of the specified button. Read/write Byte.
- `GridlineStyleTop As Byte  (read/write)`  
  Gets or sets the top gridline style of the specified button. Read/write Byte.
- `GridlineStyleRight As Byte  (read/write)`  
  Gets or sets the right gridline style of the specified button. Read/write Byte.
- `GridlineStyleBottom As Byte  (read/write)`  
  Gets or sets the bottom gridline style of the specified button. Read/write Byte.
- `GridlineWidthLeft As Byte  (read/write)`  
  Gets or sets the width of the left gridline for the specified button. Read/write Byte.
- `GridlineWidthTop As Byte  (read/write)`  
  Gets or sets the width of the top gridline for the specified button. Read/write Byte.
- `GridlineWidthRight As Byte  (read/write)`  
  Gets or sets the width of the right gridline for the specified button. Read/write Byte.
- `GridlineWidthBottom As Byte  (read/write)`  
  Gets or sets the width of the bottom gridline for the specified button. Read/write Byte.
- `GridlineColor As Long  (read/write)`  
  Gets or sets the color of the gridline for the specified button. Read/write Long.
- `HorizontalAnchor As AcHorizontalAnchor  (read/write)`  
  Gets or sets an AcHorizontalAnchor constant that indicates how the button is anchored horizontally within its layout. Read/write.
- `VerticalAnchor As AcVerticalAnchor  (read/write)`  
  Gets or sets an AcVerticalAnchor constant that indicates how the specified button is anchored vertically within its layout. Read/write.
- `LayoutID As Long  (read-only)`  
  Returns the unique identifier for the layout that contains the specified button. Read-only Long.
- `BackStyle As Byte  (read/write)`  
  Use the BackStyle property to specify whether a control will be transparent. Read/write Byte.
- `CursorOnHover As AcCursorOnHover  (read/write)`  
  Gets or sets the type of mouse cursor displayed when the user rests the pointer on a button. Read/write AcCursorOnHover.
- `PictureCaptionArrangement As AcPictureCaptionArrangement  (read/write)`  
  Gets or sets how the specified button's caption is arranged in relation to its picture. Read/write AcPictureCaptionArrangement.
- `Alignment As Byte  (read/write)`  
  Gets or sets how the specified button's caption and picture are aligned. Read/write Byte.
- `ForeThemeColorIndex As Long  (read/write)`  
  Gets or sets a value that represents a color in the applied color theme associated with the ForeColor property of the specified object. Read/write Long.
- `ForeTint As Single  (read/write)`  
  Gets or sets the tint that is applied to the theme color in the ForeColor property of the specified object. Read/write Single.
- `ForeShade As Single  (read/write)`  
  Gets or sets the shade that is applied to the theme color in the ForeColor property of the specified object. Read/write Single.
- `ThemeFontIndex As Long  (read/write)`  
  Gets or sets the font index that represents a font in the applied theme associated with the FontName property of the specified object. Read/write Long.
- `GridlineThemeColorIndex As Long  (read/write)`  
  Gets or sets the theme color index that represents a color in the applied color theme associated with the GridlineColor property of the specified object. Read/write Long.
- `GridlineTint As Single  (read/write)`  
  Gets or sets the tint applied to the theme color in the GridlineColor property of the specified object. Read/write Single.
- `GridlineShade As Single  (read/write)`  
  Gets or sets the shade applied to the theme color in the GridlineColor property of the specified object. Read/write Single.
- `NavigationTargetName As String  (read/write)`  
  Gets or sets the name of the form or report to be displayed when the user chooses the specified navigation button. Read/write String.
- `NavigationWhereClause As String  (read/write)`  
  Gets or sets the filter criteria applied when opening the form or report associated with the specified navigation button. Read/write String.
- `ParentTab As NavigationButton  (read-only)`  
  Gets the top-level navigation button associated with the specified navigation button. Read-only NavigationButton.
- `BackThemeColorIndex As Long  (read/write)`  
  Gets or sets a value that represents a color in the applied color theme associated with the BackColor property of the specified object. Read/write Long.
- `BackTint As Single  (read/write)`  
  Gets or sets the tint that is applied to the theme color in the BackColor property of the specified object. Read/write Single.
- `BackShade As Single  (read/write)`  
  Gets or sets the shade applied to the theme color in the BackColor property of the specified object. Read/write Single.
- `BorderThemeColorIndex As Long  (read/write)`  
  Gets or sets a value that represents a color in the applied color theme associated with the BorderColor property of the specified object. Read/write Long.
- `BorderTint As Single  (read/write)`  
  Gets or sets the tint that is applied to the theme color in the BorderColor property of the specified object. Read/write Single.
- `BorderShade As Single  (read/write)`  
  Gets or sets the shade applied to the theme color in the BorderColor property of the specified object. Read/write Single.
- `BorderWidth As Byte  (read/write)`  
  Use the BorderWidth property to specify the width of a control's border. Read/write Byte.
- `BorderStyle As Byte  (read/write)`  
  Specifies how a control's border appears. Read/write Byte.
- `HoverColor As Long  (read/write)`  
  Gets or sets the color displayed on the control when the user rests the mouse cursor over it. Read/write Long.
- `HoverThemeColorIndex As Long  (read/write)`  
  Gets or sets the theme color index that represents a color in the applied color theme associated with the HoverColor property of the specified object. Read/write Long.
- `HoverTint As Single  (read/write)`  
  Gets or sets the tint applied to the theme color in the HoverColor property of the specified object. Read/write Single.
- `HoverShade As Single  (read/write)`  
  Gets or sets the shade that is applied to the theme color in the HoverColor property of the specified object. Read/write Single.
- `HoverForeColor As Long  (read/write)`  
  Gets or sets the color of the text displayed on the control when the user rests the mouse cursor over it. Read/write Long.
- `HoverForeThemeColorIndex As Long  (read/write)`  
  Gets or sets the theme color index that represents a color in the applied color theme associated with the HoverForeColor property of the specified object. Read/write Long.
- `HoverForeTint As Single  (read/write)`  
  Gets or sets the tint applied to the theme color in the HoverForeColor property of the specified object. Read/write Single.
- `HoverForeShade As Single  (read/write)`  
  Gets or sets the shade applied to the theme color in the HoverForeColor property of the specified object. Read/write Single.
- `PressedColor As Long  (read/write)`  
  Gets or sets the color displayed on the control when the user chooses it. Read/write Long.
- `PressedThemeColorIndex As Long  (read/write)`  
  Gets or sets the theme color index that represents a color in the applied color theme associated with the PressedColor property of the specified object. Read/write Long.
- `PressedTint As Single  (read/write)`  
  Gets or sets the tint applied to the theme color in the PressedColor property of the specified object. Read/write Single.
- `PressedShade As Single  (read/write)`  
  Gets or sets the shade that is applied to the theme color in the PressedColor property of the specified object. Read/write Single.
- `PressedForeColor As Long  (read/write)`  
  Gets or sets the color of the text displayed on the control when the user chooses it. Read/write Long.
- `PressedForeThemeColorIndex As Long  (read/write)`  
  Gets or sets the theme color index that represents a color in the applied color theme associated with the PressedForeColor property of the specified object. Read/write Long.
- `PressedForeTint As Single  (read/write)`  
  Gets or sets the tint applied to the theme color in the PressedForeColor property of the specified object. Read/write Single.
- `PressedForeShade As Single  (read/write)`  
  Gets or sets the shade applied to the theme color in the PressedForeColor property of the specified object. Read/write Single.
- `Shape As Long  (read/write)`  
  Returns or sets the shape of the specified navigation button. Read/write Long.
- `Bevel As Long  (read/write)`  
  Gets or sets the bevel effect applied to the specified object. Read/write Long.
- `Glow As Long  (read/write)`  
  Gets or sets the glow effect applied to the specified object. Read/write Long.
- `Shadow As Long  (read/write)`  
  Gets or sets the shadow effect applied to the specified object. Read/write Long.
- `SoftEdges As Long  (read/write)`  
  Gets or sets the soft edges effect applied to the specified object. Read/write Long.
- `QuickStyle As Long  (read/write)`  
  Gets or sets the quick style that is applied to the specified object. Read/write Long.
- `QuickStyleMask As Long  (read/write)`  
  Gets or sets modifications to the quick style that is applied to the specified object. Read/write Long.
- `Gradient As Long  (read/write)`  
  Gets or sets the gradient fill applied to the specified object. Read/write Long.
- `BackColor As Long  (read/write)`  
  Gets or sets the interior color of the specified object. Read/write Long.
- `BorderColor As Long  (read/write)`  
  Use the BorderColor property to specify the color of a control's border. Read/write Long.

## Methods (4)

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

## Events (12)

- `Click()`  
  The Click event occurs when the user presses and then releases a mouse button over an object.
- `Enter()`  
  The Enter event occurs before a control actually receives the focus from a control on the same form or report.
- `Exit(Cancel As Integer)`  
  The Exit event occurs just before a control loses the focus to another control on the same form or report.
    - `Cancel As Integer` (required): Set to True to cancel the event.
- `GotFocus()`  
  The GotFocus event occurs when the specified object receives the focus.
- `LostFocus()`  
  The LostFocus event occurs when the specified object loses the focus.
- `DblClick(Cancel As Integer)`  
  The DblClick event occurs when the user presses and releases the left mouse button twice over an object within the double-click time limit of the system.
    - `Cancel As Integer` (required): The setting determines if the DblClick event occurs. Setting the Cancel argument to True (1) cancels the DblClick event.
- `MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)`  
  The MouseDown event occurs when the user presses a mouse button.
    - `Button As Integer` (required): The button that was pressed to trigger the event. If you need to test for the Button argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acLeftButton</b> The bit mask for the left mouse button.</p></li><li><p><b>acRightButton</b> The bit mask for the right mouse button.</p></li><li><p><b>acMiddleButton</b> The bit mask for the middle mouse button.</p></li></ul>
    - `Shift As Integer` (required): The state of the Shift, Ctrl, and Alt keys when the button specified by the Button argument was pressed or released. If you need to test for the Shift argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acShiftMask</b> The bit mask for the Shift key.</p></li><li><p><b>acCtrlMask</b> The bit mask for the Ctrl key.</p></li><li><p><b>acAltMask</b> The bit mask for the Alt key.</p></li></ul>
    - `X As Single` (required): The x coordinate for the current location of the mouse pointer, in twips.
    - `Y As Single` (required): The y coordinate for the current location of the mouse pointer, in twips.
- `MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)`  
  The MouseMove event occurs when the user moves the mouse.
    - `Button As Integer` (required): The button that was pressed or released when the event was triggered. If you need to test for the Button argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acLeftButton</b> The bit mask for the left mouse button.</p></li><li><p><b>acRightButton</b> The bit mask for the right mouse button.</p></li><li><p><b>acMiddleButton</b> The bit mask for the middle mouse button.</p></li></ul>
    - `Shift As Integer` (required): The state of the Shift, Ctrl, and Alt keys when the button specified by the Button argument was pressed or released. If you need to test for the Shift argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acShiftMask</b> The bit mask for the Shift key.</p></li><li><p><b>acCtrlMask</b> The bit mask for the Ctrl key.</p></li><li><p><b>acAltMask</b> The bit mask for the Alt key.</p></li></ul>
    - `X As Single` (required): The x coordinate for the current location of the mouse pointer, in twips.
    - `Y As Single` (required): The y coordinate for the current location of the mouse pointer, in twips.
- `MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)`  
  The MouseUp event occurs when the user releases a mouse button.
    - `Button As Integer` (required): The button that was released to trigger the event. If you need to test for the Button argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acLeftButton</b> The bit mask for the left mouse button.</p></li><li><p><b>acRightButton</b> The bit mask for the right mouse button.</p></li><li><p><b>acMiddleButton</b> The bit mask for the middle mouse button.</p></li></ul>
    - `Shift As Integer` (required): The state of the Shift, Ctrl, and Alt keys when the button specified by the Button argument was pressed or released. If you need to test for the Shift argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acShiftMask</b> The bit mask for the Shift key.</p></li><li><p><b>acCtrlMask</b> The bit mask for the Ctrl key.</p></li><li><p><b>acAltMask</b> The bit mask for the Alt key.</p></li></ul>
    - `X As Single` (required): The x coordinate for the current location of the mouse pointer, in twips.
    - `Y As Single` (required): The y coordinate for the current location of the mouse pointer, in twips.
- `KeyDown(KeyCode As Integer, Shift As Integer)`  
  The KeyDown event occurs when the user presses a key while a form or control has the focus. This event also occurs if you send a keystroke to a form or control by using the SendKeys action in a macro or the SendKeys statement in Visual Basic.
    - `KeyCode As Integer` (required): A key code, such as vbKeyF1 (the F1 key) or vbKeyHome (the Home key). To specify key codes, use the intrinsic constants shown in the Object Browser. You can prevent an object from receiving a keystroke by setting KeyCode to 0.
    - `Shift As Integer` (required): The state of the Shift, Ctrl, and Alt keys at the time of the event. If you need to test for the Shift argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acShiftMask</b> The bit mask for the Shift key.</p></li><li><p><b>acCtrlMask</b> The bit mask for the Ctrl key.</p></li><li><p><b>acAltMask</b> The bit mask for the Alt key.</p></li></ul>
- `KeyPress(KeyAscii As Integer)`  
  The KeyPress event occurs when the user presses and releases a key or key combination that corresponds to an ANSI code while a form or control has the focus. This event also occurs if you send an ANSI keystroke to a form or control by using the SendKeys action in a macro or the SendKeys statement in Visual Basic.
    - `KeyAscii As Integer` (required): Returns a numeric ANSI key code. The KeyAscii argument is passed by reference; changing it sends a different character to the object. Setting the KeyAscii argument to 0 cancels the keystroke so that the object doesn't recognize that a key was pressed.
- `KeyUp(KeyCode As Integer, Shift As Integer)`  
  The KeyUp event occurs when the user releases a key while a form or control has the focus. This event also occurs if you send a keystroke to a form or control by using the SendKeys action in a macro or the SendKeys statement in Visual Basic.
    - `KeyCode As Integer` (required): A key code, such as vbKeyF1 (the F1 key) or vbKeyHome (the Home key). To specify key codes, use the intrinsic constants shown in the Object Browser. You can prevent an object from receiving a keystroke by setting KeyCode to 0.
    - `Shift As Integer` (required): The state of the Shift, Ctrl, and Alt keys at the time of the event. If you need to test for the Shift argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acShiftMask</b> The bit mask for the Shift key.</p></li><li><p><b>acCtrlMask</b> The bit mask for the Ctrl key.</p></li><li><p><b>acAltMask</b> The bit mask for the Alt key.</p></li></ul>
