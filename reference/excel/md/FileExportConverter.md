# FileExportConverter

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244B3-0000-0000-C000-000000000046}  

Represents a file converter that is used to save files.

**Remarks:** You cannot create a new file converter or add one to the FileExportConverters collection. FileExportConverter objects are added during installation of Microsoft Office or by installing supplemental file converters.

**Example:**

```vba
MsgBox FileExportConverters(2).Extensions
```

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Excel application. Read-only.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified FileExportConverter object. Read-only.
- `Extensions As String  (read-only)`  
  Returns the file name extensions associated with the specified FileExportConverter object. Read-only String.
- `Description As String  (read-only)`  
  Returns the description for the file converter. Read-only String.
- `FileFormat As Long  (read-only)`  
  Returns an integer that identifies the file format associated with the specified FileExportConverter object. Read-only.
