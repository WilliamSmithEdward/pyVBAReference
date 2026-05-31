# PivotFormulas

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002441F-0000-0000-C000-000000000046}  

Represents the collection of formulas for a PivotTable report. Each formula is represented by a PivotFormula object.

**Remarks:** This object and its associated properties and methods aren't available for OLAP data sources because calculated fields and items aren't supported.

**Example:**

```vba
For Each pf in ActiveSheet.PivotTables(1).PivotFormulas
 Cells(r, 1).Value = pf.Formula
 r = r + 1
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
- `_Default As PivotFormula  (read-only)`

## Methods (3)

- `Item(Index As Variant) As PivotFormula`  
  Returns a single object from a collection.
    - `Index As Variant` (required): The name or index number for the object.
- `_NewEnum() As IUnknown`
- `Add(Formula As String, [UseStandardFormula As Variant]) As PivotFormula`  
  Creates a new PivotTable formula.
    - `Formula As String` (required): The new PivotTable formula.
    - `UseStandardFormula As Variant` (optional): A standard PivotTable formula.
