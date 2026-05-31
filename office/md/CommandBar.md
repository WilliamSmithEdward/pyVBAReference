# CommandBar

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0304-0000-0000-C000-000000000046}  

Represents a command bar in the container application. The CommandBar object is a member of the CommandBars collection.

**Example:**

```vba
foundFlag = False
For Each cb In CommandBars
    If cb.Name = "Forms" Then
        cb.Protection = msoBarNoChangeDock
        cb.Visible = True
        foundFlag = True
    End If
Next cb
If Not foundFlag Then
    MsgBox "The collection does not contain a Forms command bar."
End If
```

## Properties (20)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the CommandBar object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the CommandBar object was created. Read-only.
- `BuiltIn As Boolean  (read-only)`  
  Gets True if the specified command bar is a built-in command bar of the container application. Returns False if it is a custom command bar. Read-only.
- `Context As String  (read/write)`  
  Gets or sets a string that determines where a command bar will be saved. The string is defined and interpreted by the application. Read/write.
- `Controls As CommandBarControls  (read-only)`  
  Gets a CommandBarControls object that represents all the controls on a command bar. Read-only.
- `Enabled As Boolean  (read/write)`  
  Gets or sets a Boolean value that specifies whether the specified CommandBar is enabled. Read/write.
- `Height As Long  (read/write)`  
  Gets or sets the height of a CommandBar. Read/write.
- `Index As Long  (read-only)`  
  Gets a Long representing the index number for a CommandBar object in the collection. Read-only.
- `Left As Long  (read/write)`  
  Sets or gets the horizontal distance (in pixels) of the CommandBar from the left edge of the object relative to the screen. Read/write.
- `Name As String  (read/write)`  
  Gets the name of the built-in CommandBar object. Read-only.
- `NameLocal As String  (read/write)`  
  Gets the name of a built-in command bar as it's displayed in the language version of the container application, or returns or sets the name of a custom command bar. Read/write.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the CommandBar object. Read-only.
- `Position As MsoBarPosition  (read/write)`  
  Gets or sets an msoBarPosition constant representing the position of a command bar. Read/write.
- `RowIndex As Long  (read/write)`  
  Gets or sets the docking order of a command bar in relation to other command bars in the same docking area. Can be an integer greater than zero, or either of the following msoBarRow constants: msoBarRowFirst or msoBarRowLast. Read/write.
- `Protection As MsoBarProtection  (read/write)`  
  Gets or sets an msoBarProtection constant representing the way a command bar is protected from user customization. Read/write.
- `Top As Long  (read/write)`  
  Sets or gets the distance from the top edge of the specified command bar, to the top edge of the screen. For docked command bars, this property returns or sets the distance from the command bar to the top of the docking area. Read/write.
- `Type As MsoBarType  (read-only)`  
  Gets the type of command bar. Read-only.
- `Visible As Boolean  (read/write)`  
  Gets or sets the Visible property of the command bar. True if the command bar is visible. Read/write.
- `Width As Long  (read/write)`  
  Gets or sets the width (in pixels) of the specified command bar. Read/write.
- `AdaptiveMenu As Boolean  (read/write)`  
  Gets a Boolean value that specifies whether the command bar should include an adaptive menu. Read/write.

## Methods (4)

- `Delete()`  
  Deletes the CommandBar object from the collection.
- `FindControl([Type As Variant], [Id As Variant], [Tag As Variant], [Visible As Variant], [Recursive As Variant]) As CommandBarControl`  
  Gets a CommandBarControl object that fits a specified criteria.
    - `Type As Variant` (optional): The type of control.
    - `Id As Variant` (optional): The identifier of the control.
    - `Tag As Variant` (optional): The tag value of the control.
    - `Visible As Variant` (optional): True to include only visible command bar controls in the search. The default value is False. Visible command bars include all visible toolbars and any menus that are open at the time the FindControl method is executed.
    - `Recursive As Variant` (optional): True to include the command bar and all of its pop-up subtoolbars in the search. This argument only applies to the CommandBar object. The default value is False.
- `Reset()`  
  Resets a built-in command bar to its default configuration.
- `ShowPopup([x As Variant], [y As Variant])`  
  Displays a command bar as a shortcut menu at the specified coordinates or at the current pointer coordinates.
    - `x As Variant` (optional): The _x_-coordinate on which the location of the shortcut menu is based. If this argument is omitted, the current _x_-coordinate of the pointer is used.
    - `y As Variant` (optional): The _y_-coordinate on which the location of the shortcut menu is based. If this argument is omitted, the current _y_-coordinate of the pointer is used.
