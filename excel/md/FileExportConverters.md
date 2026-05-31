# FileExportConverters

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244B4-0000-0000-C000-000000000046}  

A collection of FileExportConverter objects that represent all the file converters available for saving files.

**Remarks:** Use the FileExportConverters property of the Application object to return the FileExportConverters collection. The Add method is not available for the FileExportConverters collection. FileExportConverter objects are added during installation of Microsoft Office or by installing supplemental converters.

**Example:**

```vba
MsgBox FileExportConverters(2).Extensions
```

## Properties (7)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Excel application. Read-only.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified FileExportConverters object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of file converters in the collection. Read-only.
- `_Default As FileExportConverter  (read-only)`
- `_NewEnum As IUnknown  (read-only)`
- `Item As FileExportConverter  (read-only)`  
  Returns an individual FileExportConverter object from a collection.
