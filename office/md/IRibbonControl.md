# IRibbonControl

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0395-0000-0000-C000-000000000046}  

Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility.

**Remarks:** The IRibbonControl object contains the name (ID) of the control and the current Window object for the Ribbon UI control.

**Example:**

```vba
public void ButtonOnAction(IRibbonControl control)
{
 MessageBox.Show("Button clicked: " + control.Id);
}

public void ToggleButtonOnAction(IRibbonControl control, bool pressed)
{
...if (pressed)
 MessageBox.Show("ToggleButton was switched on.");
 else
 MessageBox.Show("ToggleButton was switched off.");
}
```

## Properties (3)

- `Id As String  (read-only)`  
  Gets the ID of the control specified in the Ribbon XML markup customization file. Read-only.
- `Context As Object  (read-only)`  
  Represents the active window containing the Ribbon user interface that triggers a callback procedure. Read-only.
- `Tag As String  (read-only)`  
  Used to store arbitrary strings and fetch them at runtime. Read-only.
