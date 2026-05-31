# ModelColumnChanges

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244E8-0000-0000-C000-000000000046}  

A collection of ModelColumnChange objects representing columns for which the data type was changed in the Excel data model.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Excel application. Read-only.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ModelColumnChanges object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of ModelColumnChange objects in a ModelColumnChanges object. Read-only.
- `_Default As ModelColumnChange  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (1)

- `Item(Index As Variant) As ModelColumnChange`  
  Returns a single object from the ModelColumnChanges object.
    - `Index As Variant` (required): The index number for the object.
