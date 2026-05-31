# CalculatedItems

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024421-0000-0000-C000-000000000046}  

A collection of PivotItem objects that represents all the calculated items in the specified PivotTable report.

**Remarks:** A PivotTable report that contains January, February, and March items could have a calculated item named FirstQuarter defined as the sum of the amounts in January, February, and March. Use the CalculatedItems method of the PivotField object to return the CalculatedItems collection. Use CalculatedFields (_index_), where _index_ is the name or index number of the field, to return a single PivotField object from the CalculatedFields collection.

**Example:**

```vba
Set pt = Worksheets(1).PivotTables(1)
For Each ci In pt.PivotFields("Sales").CalculatedItems
 r = r + 1
 With Worksheets(2)
 .Cells(r, 1).Value = ci.Name
 .Cells(r, 2).Value = ci.Formula
 End With
Next
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

## Methods (3)

- `Item(Index As Variant) As PivotItem`  
  Returns a single object from a collection.
    - `Index As Variant` (required): The name or index number for the object.
- `_NewEnum() As IUnknown`
- `Add(Name As String, Formula As String, [UseStandardFormula As Variant]) As PivotItem`  
  Creates a new calculated item. Returns a PivotItem object.
    - `Name As String` (required): The name of the item.
    - `Formula As String` (required): The formula for the item.
    - `UseStandardFormula As Variant` (optional): False (default) for upward compatibility. True for strings contained in any arguments that are item names; will be interpreted as having been formatted in standard U.S. English instead of local settings.
