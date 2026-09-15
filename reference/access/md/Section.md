# Section

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {BC9E4355-F037-11CD-8701-00AA003F0F07}  

A form section is part of a form such as a header, footer, or detail section.

**Remarks:** You can set section properties, which are attributes of a form that affect the appearance or behavior of that section. For example, you can set the CanGrow property to specify whether the section will increase vertically to print all the data that the section contains. Section properties are set in form Design view.

## Properties (38)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Properties As Properties  (read-only)`  
  Returns a reference to a control's Properties collection object. Read-only.
- `Controls As Children  (read-only)`  
  Returns the Controls collection of a form, subform, report, or section. Read-only Controls.
- `EventProcPrefix As String  (read/write)`  
  Gets or sets the prefix portion of an event procedure name. Read/write String.
- `ForceNewPage As Byte  (read/write)`  
  Use the ForceNewPage property to specify whether form sections (detail, footer) or report sections (header, detail, footer) print on a separate page rather than on the current page. Read/write Byte.
- `NewRowOrCol As Byte  (read/write)`  
  Use the NewRowOrCol property to specify whether a section and its associated data is printed in a new row or column within a multiple-column report or multiple-column form. Read/write Byte.
- `KeepTogether As Boolean  (read/write)`  
  Use the KeepTogether property for a section to print a form or report section all on one page. For example, you might have a group of related information that you don't want printed across two pages. The KeepTogether property applies only to form and report sections (except page headers and page footers). Read/write Boolean.
- `Visible As Boolean  (read/write)`  
  Returns or sets whether the object is visible. Read/write Boolean.
- `DisplayWhen As Byte  (read/write)`  
  Use the DisplayWhen property to specify which of a form's sections you want displayed on screen and in print. Read/write Byte.
- `CanGrow As Boolean  (read/write)`  
  Gets or sets whether the specified control automatically adjusts vertically to print or preview all the data that the control contains. Read/write Boolean.
- `CanShrink As Boolean  (read/write)`  
  Gets or sets whether the specified control automatically adjusts vertically to print or preview all the data that the section or control contains. Read/write Boolean.
- `RepeatSection As Boolean  (read/write)`  
  Use the RepeatSection property to specify whether a group header is repeated on the next page or column when a group spans more than one page or column. Read/write Boolean.
- `Height As Integer  (read/write)`  
  Gets or sets the height of the specified object in twips. Read/write Integer.
- `BackColor As Long  (read/write)`  
  Gets or sets the interior color of the specified object. Read/write Long.
- `SpecialEffect As Byte  (read/write)`  
  Use the SpecialEffect property to specify whether special formatting will apply to the specified object. Read/write Byte.
- `Tag As String  (read/write)`  
  Stores extra information about a form, report, section, or control needed by a Microsoft Access application. Read/write String.
- `OnFormat As String  (read/write)`  
  Sets or returns the value of the On Format box in the Properties window of a report section. Read/write String.
- `OnPrint As String  (read/write)`  
  Sets or returns the value of the On Print box in the Properties window of a report section. Read/write String.
- `OnRetreat As String  (read/write)`  
  Sets or returns the value of the On Retreat box in the Properties window of a report section. Read/write String.
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
- `HasContinued As Boolean  (read/write)`  
  Use the HasContinued property to determine if part of the current section begins on the previous page. Read/write Boolean.
- `WillContinue As Boolean  (read/write)`  
  Determines if the current section will continue on the following page. Read/write Boolean.
- `InSelection As Boolean  (read/write)`  
  Use the InSelection property to determine or specify whether a control on a form in Design view is selected. Read/write Boolean.
- `Name As String  (read/write)`  
  Use the Name property to specify or determine the string expression that identifies the name of an object. Read/write String.
- `AutoHeight As Boolean  (read/write)`  
  Gets or sets whether a section's height is adjusted automatically when controls are resized. Read/write Boolean.
- `AlternateBackColor As Long  (read/write)`  
  Gets or sets the background color to display on alternate rows of the specified section. Read/write Long.
- `OnPaint As String  (read/write)`  
  Sets or returns the value of the On Paint box in the Properties window of a form or report. Read/write String.
- `BackThemeColorIndex As Long  (read/write)`  
  Gets or sets a value that represents a color in the applied color theme associated with the BackColor property of the specified object. Read/write Long.
- `BackTint As Single  (read/write)`  
  Gets or sets the tint that is applied to the theme color in the BackColor property of the specified object. Read/write Single.
- `BackShade As Single  (read/write)`  
  Gets or sets the shade applied to the theme color in the BackColor property of the specified object. Read/write Single.
- `AlternateBackThemeColorIndex As Long  (read/write)`  
  Gets or sets a value that represents a color in the applied color theme associated with the AlternateBackColor property of the section. Read/write Long.
- `AlternateBackTint As Single  (read/write)`  
  Gets or sets the tint applied to the theme color in the AlternateBackColor property of the section. Read/write Single.
- `AlternateBackShade As Single  (read/write)`  
  Gets or sets the shade applied to the theme color in the AlternateBackColor property of the section. Read/write Single.

## Methods (1)

- `SetTabOrder()`  
  Resets the tab order of the controls in the specified Section to their default values.

## Events (6)

- `Click()`  
  The Click event occurs when the user presses and then releases a mouse button over an object.
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
- `Paint()`  
  Occurs when the specified section is redrawn.
