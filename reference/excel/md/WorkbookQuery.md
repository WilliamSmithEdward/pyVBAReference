# WorkbookQuery

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244EB-0000-0000-C000-000000000046}  

An object that represents a query that was created by Power Query. Introduced in Office 2016.

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Name As String  (read/write)`  
  The name of the query. Read/write String.
- `Formula As String  (read/write)`  
  The Power Query M formula for the object. Read/write String.
- `Description As String  (read/write)`  
  The description of the query. Read/write String.
- `_Default As String  (read/write)`

## Methods (2)

- `Delete([DeleteConnection As Variant])`  
  Deletes this query and its underlying connection and removes it from the Queries collection.
    - `DeleteConnection As Variant` (optional): True To delete the both the query and its underlying connection . The default is False.
- `Refresh()`  
  Refreshes this query and its underlying connection.
