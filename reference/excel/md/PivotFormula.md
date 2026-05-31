# PivotFormula

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002441E-0000-0000-C000-000000000046}  

Represents a formula used to calculate results in a PivotTable report.

**Remarks:** This object and its associated properties and methods aren't available for OLAP data sources because calculated fields and items aren't supported.

**Example:**

```vba
Worksheets(1).PivotTables(1).PivotFormulas(1).Index = 2
```

## Properties (8)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `_Default As String  (read/write)`
- `Formula As String  (read/write)`  
  Returns or sets a String value that represents the object's formula in A1-style notation and in the language of the macro.
- `Value As String  (read/write)`  
  Returns or sets a String value that represents the name of the specified formula in the PivotTable formula.
- `Index As Long  (read/write)`  
  Returns or sets a Long value that represents the index number of the PivotFormula object within the PivotFormulas collection.
- `StandardFormula As String  (read/write)`  
  Returns or sets a String specifying formulas with standard English (United States) formatting. Read/write.

## Methods (1)

- `Delete()`  
  Deletes the object.
