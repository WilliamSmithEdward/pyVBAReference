# ModelTableNameChanges

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244E3-0000-0000-C000-000000000046}  

A collection of ModelTableNameChange objects representing table names before and after a table name change in the Excel data model.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Excel application. Read-only.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ModelTableNameChanges object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of ModelTableNameChange objects in a ModelTableNameChanges collection. Read-only.
- `_Default As ModelTableNameChange  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (1)

- `Item(Index As Variant) As ModelTableNameChange`  
  Returns a single ModelTableNameChange object from the ModelTableNameChanges collection.
    - `Index As Variant` (required): The index number for the object.
