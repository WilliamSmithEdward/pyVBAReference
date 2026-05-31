# ListRows

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024474-0000-0000-C000-000000000046}  

A collection of all the ListRow objects in the specified ListObject object.

**Remarks:** Each ListRow object represents a row in the table.

**Example:**

```vba
Set myNewRow = Worksheets(1).ListObject(0).ListRows.Add
```

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `_Default As ListRow  (read-only)`
- `_NewEnum As IUnknown  (read-only)`
- `Item As ListRow  (read-only)`  
  Returns a single object from a collection.
- `Count As Long  (read-only)`  
  Returns an Integer value that represents the number of objects in the collection.

## Methods (1)

- `Add([Position As Variant], [AlwaysInsert As Variant]) As ListRow`  
  Adds a new row to the table represented by the specified ListObject.
    - `Position As Variant` (optional): Integer. Specifies the relative position of the new row.
    - `AlwaysInsert As Variant` (optional): Boolean. Specifies whether to always shift data in cells below the last row of the table when the new row is inserted, regardless if the row below the table is empty. If True, the cells below the table will be shifted down one row. If False, if the row below the table is empty, the table will expand to occupy that row without shifting cells below it, but if the row below the table contains data, those cells will be shifted down when the new row is inserted.
