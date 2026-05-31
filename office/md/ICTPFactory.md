# ICTPFactory

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C033D-0000-0000-C000-000000000046}  

Used to create a custom task pane.

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

- `CreateCTP(CTPAxID As String, CTPTitle As String, [CTPParentWindow As Variant]) As _CustomTaskPane`  
  Creates an instance of a custom task pane.
    - `CTPAxID As String` (required): The CLSID or ProgID of a Microsoft ActiveX object.
    - `CTPTitle As String` (required): The title for the task pane.
    - `CTPParentWindow As Variant` (optional): The window that hosts the task pane. If not present, the parent of the task pane is the ActiveWindow of the host application.
