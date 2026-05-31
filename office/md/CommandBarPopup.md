# CommandBarPopup

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C030A-0000-0000-C000-000000000046}  

Represents a popup control on a command bar.

**Remarks:** Every pop-up control contains a CommandBar object. To return the command bar from a pop-up control, apply the CommandBar property to the CommandBarPopup object. Use Controls(_index_), where _index_ is the number of the control, to return a CommandBarPopup object. Note that the Type property of the control must be msoControlPopup, msoControlGraphicPopup, msoControlButtonPopup, msoControlSplitButtonPopup, or msoControlSplitButtonMRUPopup.

**Example:**

```vba
Set myControl = Application.CommandBars.FindControl _
(Type:=msoControlPopup, Tag:="Graphics")
```

## Properties (28)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the CommandBarPopup object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the CommandBarPopup object was created. Read-only.
- `BeginGroup As Boolean  (read/write)`  
  Gets True if the specified command bar control appears at the beginning of a group of controls on the command bar. Read/write.
- `BuiltIn As Boolean  (read-only)`  
  Is True if the specified CommandBarPopup control is a built-in command bar of the container application. Returns False if it's a custom command bar. Read-only.
- `Caption As String  (read/write)`  
  Gets or sets the caption text for a command bar control. Read/write.
- `DescriptionText As String  (read/write)`  
  Gets or sets the description for a CommandBarPopup control. Read/write.
- `Enabled As Boolean  (read/write)`  
  Is True if the CommandBarPopup is enabled. Read/write.
- `Height As Long  (read/write)`  
  Gets or sets the height of a CommandBarPopup control. Read/write.
- `HelpContextId As Long  (read/write)`  
  Gets or sets the Help context Id number for the Help topic attached to the CommandBarPopup control. Read/write.
- `HelpFile As String  (read/write)`  
  Gets or sets the file name for the Help topic attached to the CommandBarPopup control. Read/write.
- `Id As Long  (read-only)`  
  Gets the ID for a built-in CommandBarPopup control. Read-only.
- `Index As Long  (read-only)`  
  Gets a Long representing the index number for a CommandBarPopup object in the collection. Read-only.
- `Left As Long  (read-only)`  
  Gets the horizontal position of the specified CommandBarPopup control (in pixels) relative to the left edge of the screen. Returns the distance from the left side of the docking area. Read-only.
- `OLEUsage As MsoControlOLEUsage  (read/write)`  
  Gets or sets the OLE client and OLE server roles in which a CommandBarPopup control is used when two Microsoft Office applications are merged. Read/write.
- `OnAction As String  (read/write)`  
  Gets or sets the name of a Visual Basic procedure that will run when the user clicks or changes the value of a CommandBarPopup control. Read/write.
- `Parent As CommandBar  (read-only)`  
  Gets the Parent object for the CommandBarPopup object. Read-only.
- `Parameter As String  (read/write)`  
  Gets or sets a string that an application can use to execute a command from a CommandBarPopup control. Read/write.
- `Priority As Long  (read/write)`  
  Gets or sets the priority of a CommandBarPopup control. Read/write.
- `Tag As String  (read/write)`  
  Gets or sets information about the CommandBarPopup control, such as data that can be used as an argument in procedures, or information that identifies the control. Read/write.
- `TooltipText As String  (read/write)`  
  Gets or sets the text displayed in the ScreenTip of a CommandBarPopup. Read/write.
- `Top As Long  (read-only)`  
  Gets the distance (in pixels) from the top edge of the specified CommandBarPopup control to the top edge of the screen. Read-only.
- `Type As MsoControlType  (read-only)`  
  Gets the type of CommandBarPopup control. Read-only.
- `Visible As Boolean  (read/write)`  
  Gets or sets the Visible property of the CommandBarPopup control. Read/write.
- `Width As Long  (read/write)`  
  Gets or sets the width (in pixels) of the specified CommandBarPopup control. Read/write.
- `IsPriorityDropped As Boolean  (read-only)`  
  Gets True if the CommandBarPopup control is currently dropped from the menu or toolbar based on usage statistics and layout space. (Note that this is not the same as the control's visibility, as set by the Visible property). Read-only.
- `CommandBar As CommandBar  (read-only)`  
  Gets a CommandBar object that represents the menu displayed by the specified popup control. Read-only.
- `Controls As CommandBarControls  (read-only)`  
  Gets a CommandBarControls object that represents all the controls on a popup control. Read-only.
- `OLEMenuGroup As MsoOLEMenuGroup  (read/write)`  
  Gets or sets an MsoOLEMenuGroup constant that represents the menu group that the specified CommandBarPopup control belongs to when the menu groups of the OLE server are merged with the menu groups of an OLE client (that is, when an object of the container application type is embedded in another application). Read/write.

## Methods (6)

- `Copy([Bar As Variant], [Before As Variant]) As CommandBarControl`  
  Copies a CommandBarPopup control to an existing command bar.
    - `Bar As Variant` (optional): A CommandBar object that represents the destination command bar. If this argument is omitted, the control is copied to the command bar where the control already exists.
    - `Before As Variant` (optional): A number that indicates the position for the new control on the command bar. The new control will be inserted before the control at this position. If this argument is omitted, the control is copied to the end of the command bar.
- `Delete([Temporary As Variant])`  
  Deletes the CommandBarPopup object from its collection.
    - `Temporary As Variant` (optional): True to delete the control for the current session. The application will display the control again in the next session.
- `Execute()`  
  Runs the procedure or built-in command assigned to the specified CommandBarPopup control.
- `Move([Bar As Variant], [Before As Variant]) As CommandBarControl`  
  Moves the specified CommandBarPopup control to an existing command bar.
    - `Bar As Variant` (optional): A Command object that represents the destination command bar for the control. If this argument is omitted, the control is moved to the end of the command bar where the control currently resides.
    - `Before As Variant` (optional): A number that indicates the position for the control. The control is inserted before the control currently occupying this position. If this argument is omitted, the control is inserted on the same command bar.
- `Reset()`  
  Resets a built-in CommandBarPopup control to its original function and face.
- `SetFocus()`  
  Moves the keyboard focus to the specified CommandBarPopup control. If the popup is disabled or isn't visible, this method will fail.
