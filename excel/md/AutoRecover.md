# AutoRecover

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002445A-0000-0000-C000-000000000046}  

Represents the automatic recovery features of a workbook.

**Remarks:** Properties for the AutoRecover object determine the path and time interval for backing up all files. Use the AutoRecover property of the Application object to return an AutoRecover object. Use the Path property of the AutoRecover object to set the path for where the AutoRecover file will be saved.

**Example:**

```vba
Sub SetPath()

 Application.AutoRecover.Path = "C:\"

End Sub
```

## Properties (6)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Enabled As Boolean  (read/write)`  
  True if the object is enabled. Read/write Boolean.
- `Time As Long  (read/write)`  
  Sets or returns the time interval for the AutoRecover object. Permissible values are integers from 1 to 120 minutes. The default value is 10 minutes. Read/write Long.
- `Path As String  (read/write)`  
  Returns or sets a String value that represents the complete path to where Microsoft Excel will store the AutoRecover temporary files.
