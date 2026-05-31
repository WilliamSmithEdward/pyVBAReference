# CustomTaskPane

**Type:** Class  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {C5771BE5-A188-466B-AB31-00A6A32B1B1C}  

Represents a custom task pane in the container application.

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

## Properties (9)

- `Title As String  (read-only)`  
  Gets a String representing the title of a CustomTaskPane object. Read-only.
- `Application As Object  (read-only)`  
  Gets the Application object of the host application. Read-only.
- `Window As Object  (read-only)`  
  Gets the parent window object of the CustomTaskPane object. Read-only.
- `Visible As Boolean  (read/write)`  
  True if the specified CustomTaskPane object is visible. Read/write.
- `ContentControl As Object  (read-only)`  
  Gets the Microsoft ActiveX control instance displayed in the custom task pane frame. Read-only.
- `Height As Long  (read/write)`  
  Gets or sets the height of the CustomTaskPane object (in points). Read/write.
- `Width As Long  (read/write)`  
  Gets or sets the width of the task pane specified by the CustomTaskPane object. Read/write.
- `DockPosition As MsoCTPDockPosition  (read/write)`  
  Gets or sets an enumerated value specifying the docked position of a CustomTaskPane object. Read/write.
- `DockPositionRestrict As MsoCTPDockPositionRestrict  (read/write)`  
  Gets or sets an enumerated value specifying a restriction on the orientation of a CustomTaskPane object. Read/write.

## Methods (1)

- `Delete()`  
  Deletes the active custom task pane.

## Events (2)

- `VisibleStateChange(CustomTaskPaneInst As _CustomTaskPane)`  
  Occurs when the user changes the visibility of the custom task pane.
    - `CustomTaskPaneInst As _CustomTaskPane` (required): The active task pane.
- `DockPositionStateChange(CustomTaskPaneInst As _CustomTaskPane)`  
  Occurs when the user changes the docking position of the active custom task pane.
    - `CustomTaskPaneInst As _CustomTaskPane` (required): The active custom task pane.
