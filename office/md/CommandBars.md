# CommandBars

**Type:** Class  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {55F88893-7708-11D1-ACEB-006008961DA5}  

A collection of CommandBar objects that represent the command bars in the container application.

**Example:**

```vba
For Each cbar in CommandBars
    Debug.Print cbar.Name, cbar.NameLocal, cbar.Visible
Next
```

## Properties (16)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the CommandBars object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the CommandBars object was created. Read-only.
- `ActionControl As CommandBarControl  (read-only)`  
  Gets the CommandBarControl object whose OnAction property is set to the running procedure. Read-only.
- `ActiveMenuBar As CommandBar  (read-only)`  
  Gets a CommandBar object that represents the active menu bar in the container application. Read-only.
- `Count As Long  (read-only)`  
  Gets a count of the number of command bars in the host application. Read-only.
- `DisplayTooltips As Boolean  (read/write)`  
  Is True if ScreenTips are displayed whenever the user positions the pointer over command bar controls. Read/write.
- `DisplayKeysInTooltips As Boolean  (read/write)`  
  Is True if shortcut keys are displayed in the ToolTips for each command bar control. Read/write.
- `Item As CommandBar  (read-only)`  
  Gets a CommandBar object from the CommandBars collection. Read-only.
- `LargeButtons As Boolean  (read/write)`  
  Is True if the toolbar buttons displayed are larger than normal size. Read/write.
- `MenuAnimationStyle As MsoMenuAnimation  (read/write)`  
  Gets or sets a MsoMenuAnimation that represents the way a command bar is animated. Read/write.
- `_NewEnum As IUnknown  (read-only)`
- `Parent As Object  (read-only)`  
  Gets the Parent object for the CommandBars object. Read-only.
- `AdaptiveMenus As Boolean  (read/write)`  
  This property checks or unchecks the check box control for the option to show menus in Microsoft Office as full or personalized. Read/write.
- `DisplayFonts As Boolean  (read/write)`  
  Is True if the font names in the Font box are displayed in their actual fonts. Read/write.
- `DisableCustomize As Boolean  (read/write)`  
  Is True if toolbar customization is disabled. Read/write.
- `DisableAskAQuestionDropdown As Boolean  (read/write)`  
  Is True if the Answer Wizard dropdown menu is enabled. Read/write.

## Methods (13)

- `Add([Name As Variant], [Position As Variant], [MenuBar As Variant], [Temporary As Variant]) As CommandBar`  
  Creates a new command bar and adds it to the collection of command bars.
    - `Name As Variant` (optional): The name of the new command bar. If this argument is omitted, a default name is assigned to the command bar (such as Custom 1).
    - `Position As Variant` (optional): The position or type of the new command bar. Can be one of the MsoBarPosition constants.
    - `MenuBar As Variant` (optional): True to replace the active menu bar with the new command bar. The default value is False.
    - `Temporary As Variant` (optional): True to make the new command bar temporary. Command bars are deleted when the container application is closed. The default value is False.
- `FindControl([Type As Variant], [Id As Variant], [Tag As Variant], [Visible As Variant]) As CommandBarControl`  
  Gets a CommandBarControl object that fits a specified criteria.
    - `Type As Variant` (optional): The type of control.
    - `Id As Variant` (optional): The identifier of the control.
    - `Tag As Variant` (optional): The tag value of the control.
    - `Visible As Variant` (optional): True to include only visible command bar controls in the search. The default value is False. Visible command bars include all visible toolbars and any menus that are open at the time the FindControl method is executed.
- `ReleaseFocus()`  
  Releases the user interface focus from all command bars.
- `FindControls([Type As Variant], [Id As Variant], [Tag As Variant], [Visible As Variant]) As CommandBarControls`  
  Gets the CommandBarControls collection that fits the specified criteria.
    - `Type As Variant` (optional): Is one of the MsoControlType constants specifying the type of control.
    - `Id As Variant` (optional): The control's identifier.
    - `Tag As Variant` (optional): The control's tag value.
    - `Visible As Variant` (optional): True to include only visible command bar controls in the search. The default value is False.
- `ExecuteMso(idMso As String)`  
  Executes the control identified by the _idMso_ parameter.
    - `idMso As String` (required): Identifier for the control.
- `GetEnabledMso(idMso As String) As Boolean`  
  Returns True if the control identified by the _idMso_ parameter is enabled.
    - `idMso As String` (required): Identifier for the control.
- `GetVisibleMso(idMso As String) As Boolean`  
  Returns True if the control identified by the _idMso_ parameter is visible.
    - `idMso As String` (required): Identifier for the control.
- `GetPressedMso(idMso As String) As Boolean`  
  Returns a value indicating whether the toggleButton control identified by the _idMso_ parameter is pressed.
    - `idMso As String` (required): Identifier for the control.
- `GetLabelMso(idMso As String) As String`  
  Returns the label of the control identified by the _idMso_ parameter as a String.
    - `idMso As String` (required): Identifier for the control.
- `GetScreentipMso(idMso As String) As String`  
  Returns the screentip of the control identified by the _idMso_ parameter as a String.
    - `idMso As String` (required): Identifier for the control.
- `GetSupertipMso(idMso As String) As String`  
  Returns the supertip of the control identified by the _idMso_ parameter as a String.
    - `idMso As String` (required): Identifier for the control.
- `GetImageMso(idMso As String, Width As Long, Height As Long) As IPictureDisp`  
  Returns an IPictureDisp object of the control image identified by the _idMso_ parameter scaled to the dimensions specified by width and height.
    - `idMso As String` (required): Identifier for the control.
    - `Width As Long` (required): The width of the image.
    - `Height As Long` (required): The height of the image.
- `CommitRenderingTransaction(hwnd As Long)`  
  Commits the rendering transaction. Returns Nothing.
    - `hwnd As Long` (required): A handle to the window in which to commit the rendering transaction.

## Events (1)

- `OnUpdate()`  
  Occurs when any change is made to a command bar.
