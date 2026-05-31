# ModelTableColumns

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244D6-0000-0000-C000-000000000046}  

Represents a ModelTableColumn collection of single columns inside a ModelTable object.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Excel application. Read-only.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ModelTableColumns object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of ModelTableColumn objects in a ModelTableColumns object. Read-only.
- `_Default As ModelTableColumn  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (1)

- `Item(Index As Variant) As ModelTableColumn`  
  Returns a ModelTableColumn object from the ModelTableColumns collection.
    - `Index As Variant` (required): The index number for the object.
