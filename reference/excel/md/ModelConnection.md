# ModelConnection

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244D1-0000-0000-C000-000000000046}  

Contains information for the new Model Connection Type introduced in Excel 2013 to interact with the integrated data model.

**Remarks:** Use for all PivotTables connected to the model. The ModelConnection object is a new "special" workbook connection that always exists (and cannot be deleted) in workbooks that have a data model. It's created when the model is first created in a workbook. All the properties of this workbook connection are read-only. Its name is _Workbook Data Model_.

## Properties (7)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Excel application. Read-only.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ModelConnection object. Read-only.
- `CommandText As Variant  (read/write)`  
  Returns or sets the command string for the specified data source. Read/write Variant.
- `CommandType As XlCmdType  (read/write)`  
  Returns or sets one of the XlCmdType enumeration constants. Read/write.
- `ADOConnection As Object  (read-only)`  
  The ADOConnection object is used to create an open connection to a data source. Enables add-ins, such as Powerview, to create a direct connection to the engine and hence the data model. Read-only ADOConnection object.
- `CalculatedMembers As CalculatedMembers  (read-only)`  
  Returns a CalculatedMembers object that represents the calculated members in the model connection.
