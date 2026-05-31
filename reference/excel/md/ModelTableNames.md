# ModelTableNames

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244E1-0000-0000-C000-000000000046}  

A collection of table names representing tables in the Excel data model.

**Remarks:** The ModelTableNames collection is used to represent a set of tables in the Excel data model for which changes have been made during a model operation, such as a data refresh. The table names are represented as strings in the collection.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Excel application. Read-only.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ModelTableNames object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in a ModelTableNames collection. Read-only.
- `_Default As String  (read-only)`

## Methods (1)

- `Item(Index As Variant) As String`  
  Returns a single object from the ModelTableNames collection.
    - `Index As Variant` (required): The index number or name of the object.
