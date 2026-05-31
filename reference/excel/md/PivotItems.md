# PivotItems

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020877-0000-0000-C000-000000000046}  

A collection of all the PivotItem objects in a PivotTable field.

**Remarks:** The items are the individual data entries in a field category.

**Example:**

```vba
Worksheets("sheet4").Activate
With Worksheets("sheet3").PivotTables(1)
 c = 1
 For i = 1 To .PivotFields.Count
 r = 1
 Cells(r, c) = .PivotFields(i).Name
 r = r + 1
 For x = 1 To .PivotFields(i).PivotItems.Count
 Cells(r, c) = .PivotFields(i).PivotItems(x).Name
 r = r + 1
 Next
 c = c + 1
 Next
End With
```

## Properties (4)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As PivotField  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.

## Methods (3)

- `Add(Name As String)`  
  Creates a new PivotTable item.
    - `Name As String` (required): The name of the new PivotTable item.
- `Item(Index As Variant) As Object`  
  Returns a single object from a collection.
    - `Index As Variant` (required): The name or index number for the object.
- `_NewEnum() As IUnknown`
