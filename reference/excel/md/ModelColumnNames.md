# ModelColumnNames

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244E6-0000-0000-C000-000000000046}  

A collection of ModelColumnName objects representing columns of tables in the data model.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Excel application. Read-only.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ModelColumnNames object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of ModelColumnName objects in a ModelColumnNames collection. Read-only.
- `_Default As ModelColumnName  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (1)

- `Item(Index As Variant) As ModelColumnName`  
  Returns a single object from the ModelColumnNames collection.
    - `Index As Variant` (required): The index number or name for the object.
