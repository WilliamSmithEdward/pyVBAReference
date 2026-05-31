# WorksheetDataConnection

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244D2-0000-0000-C000-000000000046}  

Used to import data into the data model from data on the worksheet such as ranges and tables.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Excel application. Read-only.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified WorksheetDataConnection object. Read-only.
- `Connection As Variant  (read-only)`  
  Returns the internal connection string to the object in Excel. Read-only Variant.
- `CommandText As Variant  (read/write)`  
  Returns or sets the command string for the specified data source. Read/write Variant.
- `CommandType As XlCmdType  (read/write)`  
  Returns or sets one of the XlCmdType enumeration constants. For a WorksheetDataConnection object, this type will be set to xlCmdExcel.
