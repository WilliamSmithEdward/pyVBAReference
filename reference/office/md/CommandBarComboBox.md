# CommandBarComboBox

**Type:** Class  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {55F88897-7708-11D1-ACEB-006008961DA5}  

Represents a combo box control on a command bar.

**Remarks:** Use Controls(_index_), where _index_ is the index number of the control, to return a CommandBarComboBox object. Note that the Type property of the control must be msoControlEdit, msoControlDropdown, msoControlComboBox, msoControlButtonDropdown, msoControlSplitDropdown, msoControlOCXDropdown, msoControlGraphicCombo, or msoControlGraphicDropdown.

**Example:**

```vba
Set combo = CommandBars("Custom").Controls(2)
With combo
    .AddItem "First Item", 1
    .AddItem "Second Item", 2
    .DropDownLines = 3
    .DropDownWidth = 75
    .ListIndex = 0
End With
```

## Properties (33)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the CommandBarComboBox object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the CommandBarComboBox object was created. Read-only.
- `BeginGroup As Boolean  (read/write)`  
  Gets True if the specified command bar control appears at the beginning of a group of controls on the command bar. Read/write.
- `BuiltIn As Boolean  (read-only)`  
  Gets True if the specified command bar control is a built-in control of the container application. Returns False if it's a custom control, or if it's a built-in control whose OnAction property has been set. Read-only.
- `Caption As String  (read/write)`  
  Gets or sets the caption text for a command bar control. Read/write.
- `DescriptionText As String  (read/write)`  
  Gets or sets the description for a command bar combo box control. Read/write.
- `Enabled As Boolean  (read/write)`  
  Gets or sets a Boolean value that specifies whether the CommandBarComboBox is enabled. Read/write.
- `Height As Long  (read/write)`  
  Gets or sets the height of a CommandBarComboBox control. Read/write.
- `HelpContextId As Long  (read/write)`  
  Gets or sets the Help context Id number for the Help topic attached to the CommandBarComboBox control. Read/write.
- `HelpFile As String  (read/write)`  
  Gets or sets the file name for the Help topic attached to the CommandBarComboBox control. Read/write.
- `Id As Long  (read-only)`  
  Gets the ID for a built-in CommandBarComboBox control. Read-only.
- `Index As Long  (read-only)`  
  Gets a Long representing the index number for a CommandBarComboBox object in the collection. Read-only.
- `Left As Long  (read-only)`  
  Gets the horizontal position of the CommandBarComboBox control (in pixels) relative to the left edge of the screen. Returns the distance from the left side of the docking area. Read-only.
- `OLEUsage As MsoControlOLEUsage  (read/write)`  
  Gets or sets the OLE client and OLE server roles in which a CommandBarComboBox control will be used when two Microsoft Office applications are merged. Read/write.
- `OnAction As String  (read/write)`  
  Gets or sets the name of a Visual Basic procedure that will run when the user clicks or changes the value of a CommandBarComboBox control. Read/write.
- `Parent As CommandBar  (read-only)`  
  Gets the Parent object for the CommandBarComboBox object. Read-only.
- `Parameter As String  (read/write)`  
  Gets or sets a string that an application can use to execute a command from a CommandBarComboBox control. Read/write.
- `Priority As Long  (read/write)`  
  Gets or sets the priority of a CommandBarComboBox control. A control's priority determines whether the control can be dropped from a docked command bar if the command bar controls can't fit in a single row. Read/write.
- `Tag As String  (read/write)`  
  Gets or sets information about the CommandBarComboBox control, such as data that can be used as an argument in procedures, or information that identifies the control. Read/write.
- `TooltipText As String  (read/write)`  
  Gets or sets the text displayed in the ScreenTip of a CommandBarComboBox. Read/write.
- `Top As Long  (read-only)`  
  Gets the distance (in pixels) from the top edge of the specified CommandBarComboBox control to the top edge of the screen. Read-only.
- `Type As MsoControlType  (read-only)`  
  Gets the type of CommandBarComboBox control. Read-only.
- `Visible As Boolean  (read/write)`  
  Gets or sets the Visible property for the CommandBarComboBox control. True if the CommandBarControl is visible. Read/write.
- `Width As Long  (read/write)`  
  Gets or sets the width (in pixels) of the specified CommandBarComboBox control. Read/write.
- `IsPriorityDropped As Boolean  (read-only)`  
  Gets True if the control is currently dropped from the menu or toolbar based on usage statistics and layout space. (Note that this is not the same as the control's visibility, as set by the Visible property). Read-only.
- `DropDownLines As Long  (read/write)`  
  Gets or sets the number of lines in a command bar combo box control. The combo box control must be a custom control and it must be a drop-down list box or a combo box. Read/write.
- `DropDownWidth As Long  (read/write)`  
  Gets or sets the width (in pixels) of the list for the specified command bar combo box control. Read/write.
- `List As String  (read/write)`  
  Gets or sets an item in the CommandBarComboBox control. Read/write.
- `ListCount As Long  (read-only)`  
  Gets the number of list items in a CommandBarComboBox control. Read-only.
- `ListHeaderCount As Long  (read/write)`  
  Gets or sets the number of list items in a CommandBarComboBox control that appears above the separator line. Read/write.
- `ListIndex As Long  (read/write)`  
  Gets or sets the index number of the selected item in the list portion of the CommandBarComboBox control. If nothing is selected in the list, this property returns zero. Read/write.
- `Style As MsoComboStyle  (read/write)`  
  Gets or sets the way a CommandBarComboBox control is displayed. Can be either of the following msoComboStyle constants: msoComboLabel or msoComboNormal. Read/write.
- `Text As String  (read/write)`  
  Gets or sets the text in the display or edit portion of the CommandBarComboBox control. Read/write.

## Methods (9)

- `Copy([Bar As Variant], [Before As Variant]) As CommandBarControl`  
  Copies a command bar combo box control to an existing command bar.
    - `Bar As Variant` (optional): A CommandBar object that represents the destination command bar. If this argument is omitted, the control is copied to the command bar where the control already exists.
    - `Before As Variant` (optional): A number that indicates the position for the new control on the command bar. The new control will be inserted before the control at this position. If this argument is omitted, the control is copied to the end of the command bar.
- `Delete([Temporary As Variant])`  
  Deletes a CommandBarCombo control object from its collection.
    - `Temporary As Variant` (optional): True to delete the control for the current session. The application will display the control again in the next session.
- `Execute()`  
  Runs the procedure or built-in command assigned to the specified CommandBarComboBox control.
- `Move([Bar As Variant], [Before As Variant]) As CommandBarControl`  
  Moves the specified control to an existing command bar.
    - `Bar As Variant` (optional): A Command object that represents the destination command bar for the control. If this argument is omitted, the control is moved to the end of the command bar where the control currently resides.
    - `Before As Variant` (optional): A number that indicates the position for the control. The control is inserted before the control currently occupying this position. If this argument is omitted, the control is inserted on the same command bar.
- `Reset()`  
  Resets a built-in command bar to its default configuration, or resets a built-in CommandBarComboBox control to its original function and face.
- `SetFocus()`  
  Moves the keyboard focus to the specified CommandBarComboBox control. If the control is disabled or isn't visible, this method will fail.
- `AddItem(Text As String, [Index As Variant])`  
  Adds a list item to the specified command bar combo box control. The combo box control must be a custom control and must be a drop-down list box or a combo box.
    - `Text As String` (required): The text added to the control.
    - `Index As Variant` (optional): The position of the item in the list. If this argument is omitted, the item is added to the end of the list.
- `Clear()`  
  Removes all list items from a command bar combo box control (a drop-down list box or a combo box).
- `RemoveItem(Index As Long)`  
  Removes an item from a CommandBarComboBox control.
    - `Index As Long` (required): The item to be removed from the list.

## Events (1)

- `Change(Ctrl As CommandBarComboBox)`  
  Occurs when the end user changes the selection in a CommandBar combo box.
    - `Ctrl As CommandBarComboBox` (required): Represents a CommandBar combo box.
