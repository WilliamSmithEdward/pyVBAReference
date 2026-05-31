# ICustomTaskPaneConsumer

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C033E-0000-0000-C000-000000000046}  

An interface that provides access to the CTPFactoryAvailable method that is used to create an instance of a custom task pane.

**Example:**

```vba
public class Connect : Object, Extensibility.IDTExtensibility2, ICustomTaskPaneConsumer
...
object missing = Type.Missing;
public CustomTaskPane CTP = null;

public void CTPFactoryAvailable(ICTPFactory CTPFactoryInst)
{
 CTP = CTPFactoryInst.CreateCTP("SampleActiveX.myControl", "Task Pane Example", missing);
 sampleAX = (myControl)CTP.ContentControl;
 sampleAX.InsertTextClicked += new InsertTextEventHandler(sampleAX_InsertTextClicked);
 CTP.Visible = true;
}
...
```

## Methods (1)

- `CTPFactoryAvailable(CTPFactoryInst As ICTPFactory)`  
  Passes an ICTPFactory object to a Microsoft ActiveX add-in that can then be used when creating a custom task pane.
    - `CTPFactoryInst As ICTPFactory` (required): The object is used by an add-in to create a task pane.
