# IRibbonExtensibility

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0396-0000-0000-C000-000000000046}  

The interface through which the Ribbon user interface (UI) communicates with a COM add-in to customize the UI.

**Remarks:** The IRibbonExtensibility interface has a single method, GetCustomUI.

**Example:**

```vba
public class Connect : Object, Extensibility.IDTExtensibility2, IRibbonExtensibility
...
public string GetCustomUI(string RibbonID)
{
 StreamReader customUIReader = new System.IO.StreamReader("C:\\RibbonXSampleCS\\customUI.xml");
 string customUIData = customUIReader.ReadToEnd();
 return customUIData;
}
```

## Methods (1)

- `GetCustomUI(RibbonID As String) As String`  
  Loads the XML markup, either from an XML customization file or from XML markup embedded in the procedure, that customizes the Ribbon user interface.
    - `RibbonID As String` (required): The ID for the RibbonX UI.
