# Connections

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024486-0000-0000-C000-000000000046}  

A collection of WorkbookConnection objects for the specified workbook.

**Example:**

```vba
ActiveWorkbook.Connections.AddFromFile _
 "C:\Documents and Settings\myComputer\My Documents\My Data Sources\Northwind 2007 Customers.odc"
```

## Properties (6)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns the number of objects in the collection. Read-only Long.
- `_Default As WorkbookConnection  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (3)

- `Item(Index As Variant) As WorkbookConnection`  
  This method creates a connection item.
    - `Index As Variant` (required): Index value of the item.
- `Add2(Name As String, Description As String, ConnectionString As Variant, CommandText As Variant, [lCmdtype As Variant], [CreateModelConnection As Variant], [ImportRelationships As Variant]) As WorkbookConnection`
- `AddFromFile(Filename As String, [CreateModelConnection As Variant], [ImportRelationships As Variant]) As WorkbookConnection`  
  Adds a connection from the specified file.
    - `Filename As String` (required): Name of the file.
    - `CreateModelConnection As Variant` (optional): Specifies whether to create the connection to the model.
    - `ImportRelationships As Variant` (optional): Specifies whether to import the connection relationship.
