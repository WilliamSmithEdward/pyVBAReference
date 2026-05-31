# CommandBarButton

**Type:** Class  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {55F88891-7708-11D1-ACEB-006008961DA5}  

Represents a button control on a command bar.

**Example:**

```vba
Set c = CommandBars("Custom").Controls(2)
With c
If .Type = msoControlButton Then
    If .Style = msoButtonIcon Then
        .Style = msoButtonIconAndCaption
    Else
        .Style = msoButtonIcon
    End If
End If
End With
```

## Properties (33)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the CommandBarButton object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the CommandBarButton object was created. Read-only.
- `BeginGroup As Boolean  (read/write)`  
  Gets True if the specified command bar control appears at the beginning of a group of controls on the command bar. Read/write.
- `BuiltIn As Boolean  (read-only)`  
  Is True if the specified command bar control is a control of the container application. Returns False if it's a custom control, or if it's a built-in control whose OnAction property has been set. Read-only.
- `Caption As String  (read/write)`  
  Gets or sets the caption text for a command bar control. Read/write.
- `DescriptionText As String  (read/write)`  
  Gets or sets the description for a command bar button control. Read/write.
- `Enabled As Boolean  (read/write)`  
  True if the specified CommandBar or CommandBarControl is enabled. Read/write.
- `Height As Long  (read/write)`  
  Gets or sets the height of a command bar control. Read/write.
- `HelpContextId As Long  (read/write)`  
  Gets or sets the Help context Id number for the Help topic attached to the CommandBarButton control. Read/write.
- `HelpFile As String  (read/write)`  
  Gets or sets the file name for the Help topic attached to the CommandBarButton control. Read/write.
- `Id As Long  (read-only)`  
  Gets the ID for a built-in CommandBarButton control. Read-only.
- `Index As Long  (read-only)`  
  Gets a Long representing the index number for a CommandBarButton object in the collection. Read-only.
- `Left As Long  (read-only)`  
  Sets or gets the horizontal position of the specified CommandBarButton control (in pixels) relative to the left edge of the screen. Returns the distance from the left side of the docking area. Read-only.
- `OLEUsage As MsoControlOLEUsage  (read/write)`  
  Gets or sets the OLE client and OLE server roles in which a CommandBarButton control will be used when two Microsoft Office applications are merged. Read/write.
- `OnAction As String  (read/write)`  
  Gets or sets the name of a Visual Basic procedure that will run when the user clicks or changes the value of a CommandBarButton control. Read/write.
- `Parent As CommandBar  (read-only)`  
  Gets the Parent object for the CommandBarButton object. Read-only.
- `Parameter As String  (read/write)`  
  Gets or sets a string that an application can use to execute a command from a CommandBarButton control. Read/write.
- `Priority As Long  (read/write)`  
  Gets or sets the priority of a CommandBarButton control. A control's priority determines whether the control can be dropped from a docked command bar if the command bar controls can't fit in a single row. Controls that can't fit in a single row drop off command bars from right to left. Read/write.
- `Tag As String  (read/write)`  
  Gets or sets information about the CommandBarButton control, such as data that can be used as an argument in procedures, or information that identifies the control. Read/write.
- `TooltipText As String  (read/write)`  
  Gets or sets the text displayed in the ScreenTip of a CommandBarButton control. Read/write.
- `Top As Long  (read-only)`  
  Gets the distance (in pixels) from the top edge of the specified CommandBarButton control to the top edge of the screen. Read-only.
- `Type As MsoControlType  (read-only)`  
  Gets the type of CommandBarButton control. Read-only.
- `Visible As Boolean  (read/write)`  
  Gets or sets the Visible property of the CommandBarButton control. True if the CommandBarButton is visible. Read/write.
- `Width As Long  (read/write)`  
  Gets or sets the width (in pixels) of the specified CommandBarButton control. Read/write.
- `IsPriorityDropped As Boolean  (read-only)`  
  Gets True if the CommandBarButton control is currently dropped from the menu or toolbar based on usage statistics and layout space. (Note that this is not the same as the control's visibility, as set by the Visible property). Read-only.
- `BuiltInFace As Boolean  (read/write)`  
  Is True if the face of a command bar button control is its original built-in face. Read/write.
- `FaceId As Long  (read/write)`  
  Gets or sets the Id number for the face of a CommandBarButton control. Read/write.
- `ShortcutText As String  (read/write)`  
  Gets or sets the shortcut key text displayed next to a CommandBarButton control when the button appears on a menu, submenu, or shortcut menu. Read/write.
- `State As MsoButtonState  (read/write)`  
  Gets or sets the appearance of a CommandBarButton control. Read/write.
- `Style As MsoButtonStyle  (read/write)`  
  Gets or sets the way a CommandBarButton control is displayed. Read/write.
- `HyperlinkType As MsoCommandBarButtonHyperlinkType  (read/write)`  
  Sets or gets a msoCommandBarButtonHyperlinkType constant that represents the type of hyperlink associated with the specified command bar button. Read/write.
- `Picture As IPictureDisp  (read/write)`  
  Gets or sets an IPictureDisp object representing the image of a CommandBarButton object. Read/write.
- `Mask As IPictureDisp  (read/write)`  
  Gets or sets an IPictureDisp object representing the mask image of a CommandBarButton object. The mask image determines what parts of the button image are transparent. Read/write.

## Methods (8)

- `Copy([Bar As Variant], [Before As Variant]) As CommandBarControl`  
  Copies a command bar button control to an existing command bar.
    - `Bar As Variant` (optional): A CommandBar object that represents the destination command bar. If this argument is omitted, the control is copied to the command bar where the control already exists.
    - `Before As Variant` (optional): A number that indicates the position for the new control on the command bar. The new control will be inserted before the control at this position. If this argument is omitted, the control is copied to the end of the command bar.
- `Delete([Temporary As Variant])`  
  Deletes the CommandBarButton object from its collection.
    - `Temporary As Variant` (optional): True to delete the control for the current session. The application will display the control again in the next session.
- `Execute()`  
  Runs the procedure or built-in command assigned to the specified CommandBarButton control.
- `Move([Bar As Variant], [Before As Variant]) As CommandBarControl`  
  Moves the specified CommandBarButton control to an existing command bar.
    - `Bar As Variant` (optional): A Command object that represents the destination command bar for the control. If this argument is omitted, the control is moved to the end of the command bar where the control currently resides.
    - `Before As Variant` (optional): A number that indicates the position for the control. The control is inserted before the control currently occupying this position. If this argument is omitted, the control is inserted on the same command bar.
- `Reset()`  
  Resets a built-in CommandBarButton control to its original function and face.
- `SetFocus()`  
  Moves the keyboard focus to the specified CommandBarButton control. If the button is disabled or isn't visible, this method will fail.
- `CopyFace()`  
  Copies the face of a command bar button control to the Clipboard.
- `PasteFace()`  
  Pastes the contents of the Clipboard onto a CommandBarButton.

## Events (1)

- `Click(Ctrl As CommandBarButton, CancelDefault As Boolean)`  
  Occurs when the user clicks a CommandBarButton object.
    - `Ctrl As CommandBarButton` (required): Represents a CommandBar button.
    - `CancelDefault As Boolean` (required): Is False if the default behavior associated with the CommandBarButton controls occurs, unless it's canceled by another process or add-in.
