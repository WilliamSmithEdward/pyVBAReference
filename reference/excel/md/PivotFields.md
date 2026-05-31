# PivotFields

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020875-0000-0000-C000-000000000046}  

A collection of all the PivotField objects in a PivotTable report.

**Remarks:** In some cases, it may be easier to use one of the properties that returns a subset of the PivotTable fields. The following properties are available: - ColumnFields property - DataFields property - HiddenFields property - PageFields property - RowFields property - VisibleFields property

**Example:**

```vba
With Worksheets("sheet3").PivotTables(1)
 For i = 1 To .PivotFields.Count
 MsgBox .PivotFields(i).Name
 Next
End With
```

## Properties (4)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As PivotTable  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.

## Methods (2)

- `Item(Index As Variant) As Object`  
  Returns a single object from a collection.
    - `Index As Variant` (required): The name or index number for the object.
- `_NewEnum() As IUnknown`
