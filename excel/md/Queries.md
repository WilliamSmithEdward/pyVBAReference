# Queries

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244EC-0000-0000-C000-000000000046}  

The collection of WorkbookQuery objects introduced in Office 2016.

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns an integer that represents the number of objects in the collection.
- `_Default As WorkbookQuery  (read-only)`
- `_NewEnum As IUnknown  (read-only)`
- `FastCombine As Boolean  (read/write)`  
  True to enable the fast combine feature, as long as the workbook is open. Read/write Boolean.

## Methods (2)

- `Add(Name As String, Formula As String, [Description As Variant]) As WorkbookQuery`  
  Adds a new WorkbookQuery object to the Queries collection.
    - `Name As String` (required): The name of the query.
    - `Formula As String` (required): The Power Query M formula for the new query.
    - `Description As Variant` (optional): The description of the query.
- `Item(NameOrIndex As Variant) As WorkbookQuery`  
  Returns a single object from a collection.
    - `NameOrIndex As Variant` (required): The name or index number of the item.
