# COMAddIn

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C033A-0000-0000-C000-000000000046}  

Represents a COM add-in in the Microsoft Office host application. The COMAddIn object is a member of the COMAddIns collection.

**Example:**

```vba
MsgBox Application.COMAddIns.Item("msodraa9.ShapeSelect").Description
```

## Properties (8)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the COMAddIn object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the COMAddIn object was created. Read-only.
- `Description As String  (read/write)`  
  Gets or sets a descriptive String value for the specified COMAddin object. Read/write.
- `ProgId As String  (read-only)`  
  Gets the programmatic identifier (ProgID) for the specified COMAddIn object. Read-only.
- `Guid As String  (read-only)`  
  Gets the class identifier (CLSID) for the specified COMAddIn object. Read-only.
- `Connect As Boolean  (read/write)`  
  Gets or sets the state of the connection for the specified COMAddIn object. Read/write.
- `Object As Object  (read/write)`  
  Gets or sets an object reference. Read/write.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the COMAddIn object. Read-only.
