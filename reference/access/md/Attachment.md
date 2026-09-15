# Attachment

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {3B06E979-E47C-11CD-8701-00AA003F0F07}  

This object corresponds to an attachment control. Use an attachment control when you want to manipulate the contents fields of the attachment data type.

**Remarks:** You can attach a maximum of two gigabytes of data (the maximum size for an Access database). Individual files cannot exceed 256 megabytes in size. Office Access 2007 and later support the following graphic file formats natively, meaning the attachment control renders them without the need for additional software. - BMP (Windows Bitmap) - RLE (Run Length Encoded Bitmap) - DIB (Device Independent Bitmap) - GIF (Graphics Interchange Format) - JPEG, JPG, JPE (Joint Photographic Experts Group) - EXIF (Exchangeable File Format) - PNG (Portable Network Graphics) - TIFF, TIF (Tagged Image File Format) - ICON, ICO (Icon) - WMF (Windows Metafile) - EMF (Enhanced Metafile) As a rule, you can attach any file that was created with one of the 2007 Microsoft Office or later system programs. You can also attach log files (.log), text files (.text, .txt), and compressed .zip files. The names of your attached files can contain any Unicode character supported by the NTFS file system used in Microsoft Windows NT (NTFS). In addition, file names must conform to these guidelines: - Names must not exceed 255 characters, including the file name extensions.

## Properties (96)

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
- `EventProcPrefix As String  (read/write)`  
  Gets or sets the prefix portion of an event procedure name. Read/write String.
- `ControlType As Byte  (read/write)`  
  Use the ControlType property in Visual Basic to determine the type of control on a form or report. Read/write Byte.
- `PictureSizeMode As Byte  (read/write)`  
  Use the PictureSizeMode property to specify how a picture for an attachment control is sized. Read/write Byte.
- `PictureAlignment As Byte  (read/write)`  
  Use the PictureAlignment property to specify where a background picture will appear in the Attachment control. Read/write Byte.
- `PictureTiling As Boolean  (read/write)`  
  Use the PictureTiling property to specify whether a background picture is tiled across the entire attachment control. Read/write Boolean.
- `Visible As Boolean  (read/write)`  
  Returns or sets whether the object is visible. Read/write Boolean.
- `DisplayWhen As Byte  (read/write)`  
  Use the DisplayWhen property to specify which of a form's controls you want displayed on screen and in print. Read/write Byte.
- `Left As Integer  (read/write)`  
  Use the Left property to specify an object's location on a form or report. Read/write Integer.
- `Top As Integer  (read/write)`  
  Use the Top property to specify an object's location on a form or report. Read/write Integer.
- `Width As Integer  (read/write)`  
  Gets or sets the width of the specified object in twips. Read/write Integer.
- `Height As Integer  (read/write)`  
  Gets or sets the height of the specified object in twips. Read/write Integer.
- `BackStyle As Byte  (read/write)`  
  Use the BackStyle property to specify whether a control will be transparent. Read/write Byte.
- `BackColor As Long  (read/write)`  
  Gets or sets the interior color of the specified object. Read/write Long.
- `SpecialEffect As Byte  (read/write)`  
  Use the SpecialEffect property to specify whether special formatting will apply to the specified object. Read/write Byte.
- `BorderStyle As Byte  (read/write)`  
  Specifies how a control's border appears. Read/write Byte.
- `OldBorderStyle As Byte  (read/write)`  
  Use this property to set or return the unedited value of the BorderStyle property for a form or control. This property is useful if you need to revert to an unedited or preferred border style. Read/write Byte.
- `BorderColor As Long  (read/write)`  
  Use the BorderColor property to specify the color of a control's border. Read/write Long.
- `BorderWidth As Byte  (read/write)`  
  Use the BorderWidth property to specify the width of a control's border. Read/write Byte.
- `ControlTipText As String  (read/write)`  
  Use the ControlTipText property to specify the text that appears in a ScreenTip when you hold the mouse pointer over a control. Read/write String.
- `HelpContextId As Long  (read/write)`  
  The HelpContextID property specifies the context ID of a topic in the custom Help file specified by the HelpFile property setting. Read/write Long.
- `Section As Integer  (read/write)`  
  You can identify these controls by the section of a form or report where the control appears. Read/write Integer.
- `IsVisible As Boolean  (read/write)`  
  Use the IsVisible property to determine whether a control on a report is visible. Read/write Boolean.
- `BeforeUpdate As String  (read/write)`
- `AfterUpdate As String  (read/write)`
- `OnEnter As String  (read/write)`  
  Sets or returns the value of the On Enter box in the Properties window of the specified object. Read/write String.
- `OnExit As String  (read/write)`  
  Sets or returns the value of the On Exit box in the Properties window of the specified object. Read/write String.
- `OnDirty As String  (read/write)`  
  Sets or returns the value of the On Dirty box in the Properties window of a form or report. Read/write String.
- `OnChange As String  (read/write)`  
  Sets or returns the value of the On Change box in the Properties window of one of the objects in the Applies To list. Read/write String.
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
- `OnAttachmentCurrent As String  (read/write)`  
  Sets or gets the value of the On Attachment Current box in the Properties window of the specified Attachment. Read/write.
- `ShortcutMenuBar As String  (read/write)`  
  Use the ShortcutMenuBar property to specify the shortcut menu that appears when you right-click the specified object. Read/write String.
- `InSelection As Boolean  (read/write)`  
  Use the InSelection property to determine or specify whether a control on a form in Design view is selected. Read/write Boolean.
- `Tag As String  (read/write)`  
  Stores extra information about a control needed by a Microsoft Access application. Read/write String.
- `Name As String  (read/write)`  
  Use the Name property to specify or determine the string expression that identifies the name of an object. Read/write String.
- `DisplayAs As AcDisplayAs  (read/write)`  
  Gets or sets an AcDisplayAs constant that specifies how a file should be displayed in the Attachment control. Read/write.
- `AttachmentCount As Long  (read-only)`  
  Gets the number of documents attached to the specified Attachment control. Read-only Long.
- `CurrentAttachment As Long  (read/write)`  
  Gets the index number of the file that has the focus in the Attachment control. Read/write Long.
- `FileName As String  (read-only)`  
  Returns the full file name of the specified attachment. Read-only String.
- `FileType As String  (read-only)`  
  Returns the file extension of the specified attachment. Read-only String.
- `FileURL As String  (read-only)`  
  Returns the Uniform Resource Locator (URL) of the specified attachment. Read-only String.
- `HorizontalAnchor As AcHorizontalAnchor  (read/write)`  
  Gets or sets an AcHorizontalAnchor constant that indicates how the attachment is anchored horizontally within its layout. Read/write.
- `VerticalAnchor As AcVerticalAnchor  (read/write)`  
  Gets or sets an AcVerticalAnchor constant that indicates how the specified attachment is anchored vertically within its layout. Read/write.
- `Layout As AcLayoutType  (read-only)`  
  Returns the type of layout for the specified attachment. Read-only AcLayoutType.
- `LeftPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the attachment and its left gridline. Read/write Integer.
- `TopPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the attachment and its top gridline. Read/write Integer.
- `RightPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the attachment and its right gridline. Read/write Integer.
- `BottomPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the attachment and its bottom gridline. Read/write Integer.
- `GridlineStyleLeft As Byte  (read/write)`  
  Gets or sets the left gridline style of the specified attachment. Read/write Byte.
- `GridlineStyleTop As Byte  (read/write)`  
  Gets or sets the top gridline style of the specified attachment. Read/write Byte.
- `GridlineStyleRight As Byte  (read/write)`  
  Gets or sets the right gridline style of the specified attachment. Read/write Byte.
- `GridlineStyleBottom As Byte  (read/write)`  
  Gets or sets the bottom gridline style of the specified attachment. Read/write Byte.
- `GridlineWidthLeft As Byte  (read/write)`  
  Gets or sets the width of the left gridline for the specified attachment. Read/write Byte.
- `GridlineWidthTop As Byte  (read/write)`  
  Gets or sets the width of the top gridline for the specified attachment. Read/write Byte.
- `GridlineWidthRight As Byte  (read/write)`  
  Gets or sets the width of the right gridline for the specified attachment. Read/write Byte.
- `GridlineWidthBottom As Byte  (read/write)`  
  Gets or sets the width of the bottom gridline for the specified attachment. Read/write Byte.
- `GridlineColor As Long  (read/write)`  
  Gets or sets the color of the gridline for the specified attachment. Read/write Long.
- `DefaultPicture As String  (read/write)`  
  Gets or sets the path and file name of the graphic to be used as a background picture on an attachment control. Read/write String.
- `LayoutID As Long  (read-only)`  
  Returns the unique identifier for the layout that contains the specified attachment. Read-only Long.
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
- `ColumnWidth As Integer  (read/write)`  
  Use the ColumnWidth property to specify the width of a column in Datasheet view. Read/write Integer.
- `ColumnOrder As Integer  (read/write)`  
  Use the ColumnOrder property to specify the order of the columns in Datasheet view. Read/write Integer.
- `ColumnHidden As Boolean  (read/write)`  
  Use the ColumnHidden property to show or hide a specified column in Datasheet view. Read/write Boolean.
- `ControlSource As String  (read/write)`  
  Use the ControlSource property to specify what data appears in the specified Attachment control. Read/write String.
- `StatusBarText As String  (read/write)`  
  Use the StatusBarText property to specify the text that is displayed in the status bar when a control is selected. Read/write String.
- `TabStop As Boolean  (read/write)`  
  Use the TabStop property to specify whether you can use the Tab key to move the focus to a control in Form view. Read/write Boolean.
- `TabIndex As Integer  (read/write)`  
  Use the TabIndex property to specify a control's place in the tab order on a form. Read/write Integer.
- `Enabled As Boolean  (read/write)`  
  Use the Enabled property to set or return the status of the conditional format in the FormatCondition object. Read/write Boolean.
- `Locked As Boolean  (read/write)`  
  The Locked property specifies whether you can edit data in a control in Form view. Read/write Boolean.
- `BackThemeColorIndex As Long  (read/write)`  
  Gets or sets a value that represents a color in the applied color theme associated with the BackColor property of the specified object. Read/write Long.
- `BackTint As Single  (read/write)`  
  Gets or sets the tint that is applied to the theme color in the BackColor property of the specified object. Read/write Single.
- `BackShade As Single  (read/write)`  
  Gets or sets the shade that is applied to the theme color in the BackColor property of the specified object. Read/write Single.
- `BorderThemeColorIndex As Long  (read/write)`  
  Gets or sets a value that represents a color in the applied color theme that is associated with the BorderColor property of the specified object. Read/write Long.
- `BorderTint As Single  (read/write)`  
  Gets or sets the tint that is applied to the theme color in the BorderColor property of the specified object. Read/write Single.
- `BorderShade As Single  (read/write)`  
  Gets or sets the shade that is applied to the theme color in the BorderColor property of the specified object. Read/write Single.
- `GridlineThemeColorIndex As Long  (read/write)`  
  Gets or sets the theme color index that represents a color in the applied color theme associated with the GridlineColor property of the specified object. Read/write Long.
- `GridlineTint As Single  (read/write)`  
  Gets or sets the tint applied to the theme color in the GridlineColor property of the specified object. Read/write Single.
- `GridlineShade As Single  (read/write)`  
  Gets or sets the shade applied to the theme color in the GridlineColor property of the specified object. Read/write Single.
- `DefaultPictureType As Byte  (read/write)`  
  Gets or sets the method used to store the image specified by the DefaultPicture property in the database. Read/write Byte.

## Methods (6)

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
- `Forward()`  
  Moves the focus to the next file in the specified Attachment control.
- `Back()`  
  Moves the focus to the previous file in the specified Attachment control.

## Events (17)

- `BeforeUpdate(Cancel As Integer)`
- `AfterUpdate()`
- `Enter()`  
  The Enter event occurs before a control actually receives the focus from a control on the same form.
- `Exit(Cancel As Integer)`  
  The Exit event occurs just before a control loses the focus to another control on the same form.
    - `Cancel As Integer` (required): The setting determines if the Exit event occurs. Setting the _Cancel_ argument to True (1) cancels the Exit event.
- `Dirty(Cancel As Integer)`  
  The Dirty event occurs when the content of the specified control changes.
    - `Cancel As Integer` (required): The setting determines if the Dirty event occurs. Setting the _Cancel_ argument to True (1) cancels the Dirty event.
- `Change()`  
  The Change event occurs when the contents of the specified control change.
- `GotFocus()`  
  The GotFocus event occurs when the specified object receives the focus.
- `LostFocus()`  
  The LostFocus event occurs when the specified object receives the focus.
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
    - `Button As Integer` (required): The button that was released to trigger the event. If you need to test for the _Button_ argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p>acLeftButton The bit mask for the left mouse button.</p></li><li><p>acRightButton The bit mask for the right mouse button.</p></li><li><p>acMiddleButton The bit mask for the middle mouse button.</p></li></ul>
    - `Shift As Integer` (required): The state of the Shift, Ctrl, and Alt keys when the button specified by the _Button_ argument was pressed or released. If you need to test for the _Shift_ argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p>acShiftMask The bit mask for the Shift key.</p></li><li><p>acCtrlMask The bit mask for the Ctrl key.</p></li><li><p>acAltMask The bit mask for the Alt key.</p></li></ul>
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
    - `KeyCode As Integer` (required): A key code, such as vbKeyF1 (the F1 key) or vbKeyHome (the Home key). To specify key codes, use the intrinsic constants shown in the Object Browser. Prevent an object from receiving a keystroke by setting _KeyCode_ to 0.
    - `Shift As Integer` (required): The state of the Shift, Ctrl, and Alt keys at the time of the event. If you need to test for the _Shift_ argument, use one of the following intrinsic constants as bit masks:<ul><li><p>acShiftMask The bit mask for the Shift key.</p></li><li><p>acCtrlMask The bit mask for the Ctrl key.</p></li><li><p>acAltMask The bit mask for the Alt key.</p></li></ul>
- `AttachmentCurrent()`  
  Occurs when the focus moves to or from a file contained in the specified Attachment control, making it the current attachment, or when the form is refreshed or requeried.
