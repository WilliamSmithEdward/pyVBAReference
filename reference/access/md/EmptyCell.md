# EmptyCell

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {3B06E986-E47C-11CD-8701-00AA003F0F07}  

Represents an empty cell in a control layout.

**Remarks:** An EmptyCell object is created when a control layout contains an empty cell. You cannot explicitly create or delete an EmptyCell object. When you type inside an empty cell, Microsoft Access automatically inserts a Label control.

## Properties (45)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Properties As Properties  (read-only)`  
  Returns a reference to a control's Properties collection object. Read-only.
- `EventProcPrefix As String  (read/write)`  
  Gets or sets the prefix portion of an event procedure name. Read/write String.
- `ControlType As Byte  (read/write)`  
  Use the ControlType property in Visual Basic to determine the type of control on a form or report. Read/write Byte.
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
- `HelpContextId As Long  (read/write)`  
  The HelpContextID property specifies the context ID of a topic in the custom Help file specified by the HelpFile property setting. Read/write Long.
- `Section As Integer  (read/write)`  
  You can identify these controls by the section of a form or report where the control appears. Read/write Integer.
- `IsVisible As Boolean  (read/write)`  
  Use the IsVisible property to determine whether a control on a report is visible. Read/write Boolean.
- `ShortcutMenuBar As String  (read/write)`  
  Use the ShortcutMenuBar property to specify the shortcut menu that appears when you right-click the specified object. Read/write String.
- `InSelection As Boolean  (read/write)`  
  Use the InSelection property to determine or specify whether a control on a form in Design view is selected. Read/write Boolean.
- `Tag As String  (read/write)`  
  Stores extra information about a form, report, section, or control needed by a Microsoft Access application. Read/write String.
- `Name As String  (read/write)`  
  Use the Name property to specify or determine the string expression that identifies the name of an object. Read/write String.
- `HorizontalAnchor As AcHorizontalAnchor  (read/write)`  
  Gets or sets an AcHorizontalAnchor constant that indicates how the cell is anchored horizontally within its layout. Read/write.
- `VerticalAnchor As AcVerticalAnchor  (read/write)`  
  Gets or sets an AcVerticalAnchor constant that indicates how the specified cell is anchored vertically within its layout. Read/write.
- `Layout As AcLayoutType  (read-only)`  
  Returns the type of layout for the specified cell. Read-only AcLayoutType.
- `LeftPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the cell and its left gridline. Read/write Integer.
- `TopPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the cell and its top gridline. Read/write Integer.
- `RightPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the cell and its right gridline. Read/write Integer.
- `BottomPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the cell and its bottom gridline. Read/write Integer.
- `GridlineStyleLeft As Byte  (read/write)`  
  Gets or sets the left gridline style of the specified cell. Read/write Byte.
- `GridlineStyleTop As Byte  (read/write)`  
  Gets or sets the top gridline style of the specified cell. Read/write Byte.
- `GridlineStyleRight As Byte  (read/write)`  
  Gets or sets the right gridline style of the specified cell. Read/write Byte.
- `GridlineStyleBottom As Byte  (read/write)`  
  Gets or sets the bottom gridline style of the specified cell. Read/write Byte.
- `GridlineWidthLeft As Byte  (read/write)`  
  Gets or sets the width of the left gridline for the specified cell. Read/write Byte.
- `GridlineWidthTop As Byte  (read/write)`  
  Gets or sets the width of the top gridline for the specified cell. Read/write Byte.
- `GridlineWidthRight As Byte  (read/write)`  
  Gets or sets the width of the right gridline for the specified cell. Read/write Byte.
- `GridlineWidthBottom As Byte  (read/write)`  
  Gets or sets the width of the bottom gridline for the specified cell. Read/write Byte.
- `GridlineColor As Long  (read/write)`  
  Gets or sets the color of the gridline for the specified cell. Read/write Long.
- `LayoutID As Long  (read-only)`  
  Returns the unique identifier for the layout that contains the specified cell. Read-only Long.
- `StatusBarText As String  (read/write)`  
  Use the StatusBarText property to specify the text that is displayed in the status bar when a control is selected. Read/write String.
- `BackThemeColorIndex As Long  (read/write)`  
  Gets or sets a value that represents a color in the applied color theme associated with the BackColor property of the specified object. Read/write Long.
- `BackTint As Single  (read/write)`  
  Gets or sets the tint that is applied to the theme color in the BackColor property of the specified object. Read/write Single.
- `BackShade As Single  (read/write)`  
  Gets or sets the shade applied to the theme color in the BackColor property of the specified object. Read/write Single.
- `GridlineThemeColorIndex As Long  (read/write)`  
  Gets or sets the theme color index that represents a color in the applied color theme associated with the GridlineColor property of the specified object. Read/write Long.
- `GridlineTint As Single  (read/write)`  
  Gets or sets the tint applied to the theme color in the GridlineColor property of the specified object. Read/write Single.
- `GridlineShade As Single  (read/write)`  
  Gets or sets the shade applied to the theme color in the GridlineColor property of the specified object. Read/write Single.

## Methods (2)

- `SizeToFit()`  
  Use the SizeToFit method to size a control so that it fits the text or image that it contains.
- `Move(Left As Variant, [Top As Variant], [Width As Variant], [Height As Variant])`  
  Moves the specified object to the coordinates specified by the argument values.
    - `Left As Variant` (required): The screen position in twips for the left edge of the object relative to the left edge of the Microsoft Access window.
    - `Top As Variant` (optional): The screen position in twips for the top edge of the object relative to the top edge of the Access window.
    - `Width As Variant` (optional): The desired width of the object in twips.
    - `Height As Variant` (optional): The desired height of the object in twips.
