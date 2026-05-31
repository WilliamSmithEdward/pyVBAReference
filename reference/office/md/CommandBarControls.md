# CommandBarControls

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0306-0000-0000-C000-000000000046}  

A collection of CommandBarControl objects that represent the command bar controls on a command bar.

**Example:**

```vba
For Each ctl In CommandBars("Standard").Controls
    ctl.Caption = CStr(ctl.Id)
Next ctl
```

## Properties (6)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the CommandBarControls object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the CommandBarControls object was created. Read-only.
- `Count As Long  (read-only)`  
  Gets a count of the numbers of controls on a command bar. Read-only.
- `Item As CommandBarControl  (read-only)`  
  Gets a CommandBarControl object from the CommandBarControls collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`
- `Parent As CommandBar  (read-only)`  
  Gets the Parent object for the CommandBarControls object. Read-only.

## Methods (1)

- `Add([Type As Variant], [Id As Variant], [Parameter As Variant], [Before As Variant], [Temporary As Variant]) As CommandBarControl`  
  Creates a new CommandBarControl object and adds it to the collection of controls on the specified command bar.
    - `Type As Variant` (optional): The type of control to be added to the specified command bar. Can be one of the following MsoControl constants: msoControlButton, msoControlEdit, msoControlDropdown, msoControlComboBox, or msoControlPopup.
    - `Id As Variant` (optional): An integer that specifies a built-in control. If the value of this argument is 1, or if this argument is omitted, a blank custom control of the specified type will be added to the command bar.
    - `Parameter As Variant` (optional): For built-in controls, this argument is used by the container application to run the command. For custom controls, you can use this argument to send information to Visual Basic procedures, or you can use it to store information about the control (similar to a second Tag property value).
    - `Before As Variant` (optional): A number that indicates the position of the new control on the command bar. The new control will be inserted before the control at this position. If this argument is omitted, the control is added at the end of the specified command bar.
    - `Temporary As Variant` (optional): True to make the new control temporary. Controls are automatically deleted when the container application is closed. The default value is False.
