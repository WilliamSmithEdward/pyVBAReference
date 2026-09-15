# Image

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {3B06E94D-E47C-11CD-8701-00AA003F0F07}  

This object corresponds to an image control. The image control can add a picture to a form or report. For example, you could include an image control for a logo on an Invoice report.

**Remarks:** Use the image control or an Unbound object frame for unbound pictures. The advantage of using the image control is that it's faster to display. The advantage of using the unbound object frame is that you can edit the object directly from the form or report.

## Properties (71)

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
- `Picture As String  (read/write)`  
  Use the Picture property to specify a bitmap or other type of graphic to be displayed on the specified control. Read/write String.
- `PictureType As Byte  (read/write)`  
  Use the PictureType property to specify whether Microsoft Access stores an object's picture as a linked or an embedded object. Read/write Byte.
- `PictureData As Variant  (read/write)`  
  Use the PictureData property to copy the picture to another object that supports the Picture property. Read/write Variant.
- `SizeMode As Byte  (read/write)`  
  Use the SizeMode property to specify how to size a picture or other object in a bound object frame, an unbound object frame, or an image control.
- `PictureAlignment As Byte  (read/write)`  
  Use the PictureAlignment property to specify where a background picture will appear in an image control or on a form or report. Read/write Byte.
- `PictureTiling As Boolean  (read/write)`  
  Use the PictureTiling property to specify whether a background picture is tiled across the entire image control, Form window, form, or page of a report. Read/write Boolean.
- `HyperlinkAddress As String  (read/write)`  
  Use the HyperlinkAddress property to specify or determine the path to an object, document, webpage, or other destination for a hyperlink associated with the specified control. Read/write String.
- `HyperlinkSubAddress As String  (read/write)`  
  Use the HyperlinkSubAddress property to specify or determine a location within the target document specified by the HyperlinkAddress property. Read/write String.
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
- `ObjectPalette As Variant  (read/write)`  
  The ObjectPalette property specifies the palette in the application used to create a bitmap or other graphic that is loaded into the specified control by using the Picture property. Read/write Variant.
- `IsVisible As Boolean  (read/write)`  
  Use the IsVisible property to determine whether a control on a report is visible. Read/write Boolean.
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
- `ShortcutMenuBar As String  (read/write)`  
  Use the ShortcutMenuBar property to specify the shortcut menu that appears when you right-click the specified object. Read/write String.
- `ImageHeight As Long  (read/write)`  
  Use the ImageHeight property to determine the height in twips of the picture in an image control. Read/write Long.
- `ImageWidth As Long  (read/write)`  
  Use the ImageWidth property to determine the width in twips of a picture in an image control. Read/write Long.
- `InSelection As Boolean  (read/write)`  
  Use the InSelection property to determine or specify whether a control on a form in Design view is selected. Read/write Boolean.
- `Tag As String  (read/write)`  
  Stores extra information about a form, report, section, or control needed by a Microsoft Access application. Read/write String.
- `Name As String  (read/write)`  
  Use the Name property to specify or determine the string expression that identifies the name of an object. Read/write String.
- `Layout As AcLayoutType  (read-only)`  
  Returns the type of layout for the specified image control. Read-only AcLayoutType.
- `LeftPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the image control and its left gridline. Read/write Integer.
- `TopPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the image control and its top gridline. Read/write Integer.
- `RightPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the image control and its right gridline. Read/write Integer.
- `BottomPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the image control and its bottom gridline. Read/write Integer.
- `GridlineStyleLeft As Byte  (read/write)`  
  Gets or sets the left gridline style of the specified image control. Read/write Byte.
- `GridlineStyleTop As Byte  (read/write)`  
  Gets or sets the top gridline style of the specified image control. Read/write Byte.
- `GridlineStyleRight As Byte  (read/write)`  
  Gets or sets the right gridline style of the specified image control. Read/write Byte.
- `GridlineStyleBottom As Byte  (read/write)`  
  Gets or sets the bottom gridline style of the specified image control. Read/write Byte.
- `GridlineWidthLeft As Byte  (read/write)`  
  Gets or sets the width of the left gridline for the specified image control. Read/write Byte.
- `GridlineWidthTop As Byte  (read/write)`  
  Gets or sets the width of the top gridline for the specified image control. Read/write Byte.
- `GridlineWidthRight As Byte  (read/write)`  
  Gets or sets the width of the right gridline for the specified image control. Read/write Byte.
- `GridlineWidthBottom As Byte  (read/write)`  
  Gets or sets the width of the bottom gridline for the specified image control. Read/write Byte.
- `GridlineColor As Long  (read/write)`  
  Gets or sets the color of the gridline for the specified image control. Read/write Long.
- `HorizontalAnchor As AcHorizontalAnchor  (read/write)`  
  Gets or sets an AcHorizontalAnchor constant that indicates how the image is anchored horizontally within its layout. Read/write.
- `VerticalAnchor As AcVerticalAnchor  (read/write)`  
  Gets or sets an AcVerticalAnchor constant that indicates how the specified image is anchored vertically within its layout. Read/write.
- `LayoutID As Long  (read-only)`  
  Returns the unique identifier for the layout that contains the specified image control. Read-only Long.
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
  Gets or sets the shade that is applied to the theme color in the BorderColor property of the specified object. Read/write Single.
- `GridlineThemeColorIndex As Long  (read/write)`  
  Gets or sets the theme color index that represents a color in the applied color theme associated with the GridlineColor property of the specified object. Read/write Long.
- `GridlineTint As Single  (read/write)`  
  Gets or sets the tint applied to the theme color in the GridlineColor property of the specified object. Read/write Single.
- `GridlineShade As Single  (read/write)`  
  Gets or sets the shade applied to the theme color in the GridlineColor property of the specified object. Read/write Single.

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

## Events (5)

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
