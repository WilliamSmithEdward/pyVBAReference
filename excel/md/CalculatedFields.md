# CalculatedFields

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024420-0000-0000-C000-000000000046}  

A collection of PivotField objects that represents all the calculated fields in the specified PivotTable report.

**Remarks:** A report that contains Revenue and Expense fields could have a calculated field named Profit defined as the amount in the Revenue field minus the amount in the Expense field. For OLAP data sources, you cannot set this collection, and it always returns Nothing. Use the CalculatedFields method of the PivotTable object to return the CalculatedFields collection. Use CalculatedFields (_index_), where _index_ is the specified field's name or index number, to return a single PivotField object from the CalculatedFields collection.

**Example:**

```vba
For Each fld in _
 Worksheets(1).PivotTables("Pivot1").CalculatedFields
 fld.Delete
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
- `_Default As PivotField  (read-only)`

## Methods (3)

- `Item(Index As Variant) As PivotField`  
  Returns a single object from a collection.
    - `Index As Variant` (required): The name or index number for the object.
- `_NewEnum() As IUnknown`
- `Add(Name As String, Formula As String, [UseStandardFormula As Variant]) As PivotField`  
  Creates a new calculated field. Returns a PivotField object.
    - `Name As String` (required): The name of the field.
    - `Formula As String` (required): The formula for the field.
    - `UseStandardFormula As Variant` (optional): False (default) for upward compatibility. True for strings contained in any arguments that are field names; will be interpreted as having been formatted in standard U.S. English instead of local settings.
