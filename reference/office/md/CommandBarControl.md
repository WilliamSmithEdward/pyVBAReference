# CommandBarControl

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0308-0000-0000-C000-000000000046}  

Represents a command bar control. The CommandBarControl object is a member of the CommandBarControls collection. The properties and methods of the CommandBarControl object are all shared by the CommandBarButton, CommandBarComboBox, and CommandBarPopup objects.

**Remarks:** When writing Visual Basic code to work with custom command bar controls, you use the CommandBarButton, CommandBarComboBox, and CommandBarPopup objects. When writing code to work with built-in controls in the container application that cannot be represented by one of those three objects, you use the CommandBarControl object. Use Controls (_index_), where _index_ is the index number of a control, to return a CommandBarControl object. (The Type property of the control must be msoControlLabel, msoControlExpandingGrid, msoControlSplitExpandingGrid, msoControlGrid, or msoControlGauge). Variables declared as CommandBarControl can be assigned CommandBarButton, CommandBarComboBox, and CommandBarPopup values.

**Example:**

```vba
Set lbl = CommandBars.FindControl(Type:= msoControlGauge)
If lbl Is Nothing Then
    MsgBox "A control of type msoControlGauge was not found."
Else
    MsgBox "Control " & lbl.Index & " on command bar " _
        & lbl.Parent.Name & " is type msoControlGauge"
End If
```

## Properties (25)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the CommandBarControl object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the CommandBarControl object was created. Read-only.
- `BeginGroup As Boolean  (read/write)`  
  Gets True if the specified command bar control appears at the beginning of a group of controls on the command bar. Read/write.
- `BuiltIn As Boolean  (read-only)`  
  Gets True if the specified command bar control is a built-in control of the container application. Returns False if it's a custom control, or if it's a built-in control whose OnAction property has been set. Read-only.
- `Caption As String  (read/write)`  
  Gets or sets the caption text for a command bar control. Read/write.
- `DescriptionText As String  (read/write)`  
  Gets or sets the description for a command bar control. Read/write.
- `Enabled As Boolean  (read/write)`  
  Gets or sets a Boolean value specifying if the CommandBarControl is enabled. Read/write.
- `Height As Long  (read/write)`  
  Gets or sets the height of a CommandBarControl control. Read/write.
- `HelpContextId As Long  (read/write)`  
  Gets or sets the Help context Id number for the Help topic attached to the CommandBarControl. Read/write.
- `HelpFile As String  (read/write)`  
  Gets or sets the file name for the Help topic attached to the CommandBarControl. Read/write.
- `Id As Long  (read-only)`  
  Gets the ID for a built-in CommandBarControl. Read-only.
- `Index As Long  (read-only)`  
  Gets a Long representing the index number for a CommandBarControl object in the collection. Read-only.
- `Left As Long  (read-only)`  
  Gets the horizontal position of the specified CommandBarControl (in pixels) relative to the left edge of the screen. Returns the distance from the left side of the docking area. Read-only.
- `OLEUsage As MsoControlOLEUsage  (read/write)`  
  Gets or sets the OLE client and OLE server roles in which a CommandBarControl will be used when two Microsoft Office applications are merged. Read/write.
- `OnAction As String  (read/write)`  
  Gets or sets the name of a Visual Basic procedure that will run when the user clicks or changes the value of a CommandBarControl. Read/write.
- `Parent As CommandBar  (read-only)`  
  Gets the Parent object for the CommandBarControl object. Read-only.
- `Parameter As String  (read/write)`  
  Gets or sets a string that an application can use to execute a command from a CommandBarControl. Read/write.
- `Priority As Long  (read/write)`  
  Gets or sets the priority of a CommandBarControl. A control's priority determines whether the control can be dropped from a docked command bar if the command bar controls can't fit in a single row. Controls that can't fit in a single row drop off command bars from right to left. Read/write.
- `Tag As String  (read/write)`  
  Gets or sets information about the CommandBarControl, such as data that can be used as an argument in procedures, or information that identifies the control. Read/write.
- `TooltipText As String  (read/write)`  
  Gets or sets the text displayed in the ScreenTip of a CommandBarControl. Read/write.
- `Top As Long  (read-only)`  
  Gets the distance (in pixels) from the top edge of the specified CommandBarControl to the top edge of the screen. Read-only.
- `Type As MsoControlType  (read-only)`  
  Gets the type of CommandBarControl. Read-only.
- `Visible As Boolean  (read/write)`  
  Gets or sets the Visible property of the CommandBarControl. True if the CommandBarControl is visible. Read/write.
- `Width As Long  (read/write)`  
  Gets or sets the width (in pixels) of the specified CommandBarControl. Read/write.
- `IsPriorityDropped As Boolean  (read-only)`  
  Gets True if the control is currently dropped from the menu or toolbar based on usage statistics and layout space. (Note that this is not the same as the control's visibility, as set by the Visible property). Read-only.

## Methods (6)

- `Copy([Bar As Variant], [Before As Variant]) As CommandBarControl`  
  Copies a command bar control to an existing command bar.
    - `Bar As Variant` (optional): A CommandBar object that represents the destination command bar. If this argument is omitted, the control is copied to the command bar where the control already exists.
    - `Before As Variant` (optional): A number that indicates the position for the new control on the command bar. The new control will be inserted before the control at this position. If this argument is omitted, the control is copied to the end of the command bar.
- `Delete([Temporary As Variant])`  
  Deletes the CommandBarControl object from its collection.
    - `Temporary As Variant` (optional): True to delete the control for the current session. The application will display the control again in the next session.
- `Execute()`  
  Runs the procedure or built-in command assigned to the specified CommandBarControl control.
- `Move([Bar As Variant], [Before As Variant]) As CommandBarControl`  
  Moves the specified CommandBarControl to an existing command bar.
    - `Bar As Variant` (optional): A Command object that represents the destination command bar for the control. If this argument is omitted, the control is moved to the end of the command bar where the control currently resides.
    - `Before As Variant` (optional): A number that indicates the position for the control. The control is inserted before the control currently occupying this position. If this argument is omitted, the control is inserted on the same command bar.
- `Reset()`  
  Resets a built-in command bar to its default configuration, or resets a built-in CommandBarControl to its original function and face.
- `SetFocus()`  
  Moves the keyboard focus to the specified CommandBarControl. If the control is disabled or isn't visible, this method will fail.
