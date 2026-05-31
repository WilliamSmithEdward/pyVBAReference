# ModelMeasureNames

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244EA-0000-0000-C000-000000000046}  

A collection of ModelMeasureName objects representing names of measures in the Excel data model.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Excel application. Read-only.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ModelMeasureNames object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of ModelMeasureName objects in a ModelMeasureNames object. Read-only.
- `_Default As ModelMeasureName  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (1)

- `Item(Index As Variant) As ModelMeasureName`  
  Returns a single object from the ModelMeasureNames object.
    - `Index As Variant` (required): The index number for the object.
