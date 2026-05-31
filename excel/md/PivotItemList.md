# PivotItemList

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024468-0000-0000-C000-000000000046}  

A collection of all the PivotItem objects in the specified PivotTable.

**Remarks:** Each PivotItem represents an item in a PivotTable field. Use the RowItems or ColumnItems property of the PivotCell object to return a PivotItemList collection.

**Example:**

```vba
Sub CheckPivotItemList()

 ' Identify contents associated with PivotItemList.
 MsgBox "Contents associated with cell B5: " & _
 Application.Range("B5").PivotCell.RowItems.Item(1)

End Sub
```

## Properties (5)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.
- `_Default As PivotItem  (read-only)`

## Methods (2)

- `Item(Index As Variant) As PivotItem`  
  Returns a single object from a collection.
    - `Index As Variant` (required): The name or index number for the object.
- `_NewEnum() As IUnknown`
