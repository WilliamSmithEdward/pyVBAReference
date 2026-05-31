# IRibbonUI

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03A7-0000-0000-C000-000000000046}  

The object that is returned by the onLoad procedure specified on the customUI tag. The object contains methods for invalidating control properties and for refreshing the user interface.

**Remarks:** You can customize the Ribbon user interface (UI) by using callback procedures in COM add-ins. When the host application starts, the onLoad callback procedure is called. The callback procedure then returns a IRibbonUI object pointing to the user interface (UI). Use that object to invoke the Invalidate, InvalidateControl, and Refresh methods.

**Example:**

```vba
<customUI ... OnLoad="MyAddInInitialize" ...>
```

## Methods (6)

- `Invalidate()`  
  Invalidates the cached values for all of the controls of the Ribbon user interface.
- `InvalidateControl(ControlID As String)`  
  Invalidates the cached value for a single control on the Ribbon user interface.
- `InvalidateControlMso(ControlID As String)`  
  Used to invalidate a built-in control.
- `ActivateTab(ControlID As String)`  
  Activates the specified custom tab. This method returns S_FALSE if there is no Ribbon or the Ribbon is collapsed.
    - `ControlID As String` (required): Specifies the Id of the custom Ribbon tab to be activated.
- `ActivateTabMso(ControlID As String)`  
  Activates the specified built-in tab.
    - `ControlID As String` (required): Specifies the Id of the custom Ribbon tab to be activated.
- `ActivateTabQ(ControlID As String, Namespace As String)`  
  Activates the specified custom tab on the Microsoft Office Fluent Ribbon UI. Uses the fully qualified name of the tab, which includes the ID and the namespace of the tab.
    - `ControlID As String` (required): Specifies the Id of the custom Ribbon tab to be activated.
    - `Namespace As String` (required): Specifies the namespace of the tab element.
