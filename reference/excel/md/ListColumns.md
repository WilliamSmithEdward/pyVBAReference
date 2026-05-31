# ListColumns

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024472-0000-0000-C000-000000000046}  

A collection of all the ListColumn objects in the specified ListObject object.

**Remarks:** Each ListColumn object represents a column in the table.

**Example:**

```vba
Set myNewColumn = Worksheets(1).ListObject(1).ListColumns.Add
```

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `_Default As ListColumn  (read-only)`
- `_NewEnum As IUnknown  (read-only)`
- `Item As ListColumn  (read-only)`  
  Returns a single object from a collection.
- `Count As Long  (read-only)`  
  Returns an Integer value that represents the number of objects in the collection.

## Methods (1)

- `Add([Position As Variant]) As ListColumn`  
  Adds a new column to the list object.
    - `Position As Variant` (optional): Integer. Specifies the relative position of the new column that starts at 1. The previous column at this position is shifted outward.
