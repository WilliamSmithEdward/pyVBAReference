# PivotTables

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020873-0000-0000-C000-000000000046}  

A collection of all the PivotTable objects in the specified workbook.

**Remarks:** Because PivotTable report programming can be complex, it's generally easiest to record PivotTable report actions and then revise the recorded code.

**Example:**

```vba
MsgBox Worksheets("sheet3").PivotTables.Count
```

## Properties (4)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.

## Methods (3)

- `Item(Index As Variant) As PivotTable`  
  Returns a single object from a collection.
    - `Index As Variant` (required): The name or index number for the object.
- `_NewEnum() As IUnknown`
- `Add(PivotCache As PivotCache, TableDestination As Variant, [TableName As Variant], [ReadData As Variant], [DefaultVersion As Variant]) As PivotTable`  
  Adds a new PivotTable report. Returns a PivotTable object.
    - `PivotCache As PivotCache` (required): The PivotTable cache on which the new PivotTable report is based. The cache provides data for the report.
    - `TableDestination As Variant` (required): The cell in the upper-left corner of the PivotTable report's destination range (the range on the worksheet where the resulting report will be placed). You must specify a destination range on the worksheet that contains the PivotTables object specified by _expression_.
    - `TableName As Variant` (optional): The name of the new PivotTable report.
    - `ReadData As Variant` (optional): True to create a PivotTable cache that contains all records from the external database; this cache can be very large. False to enable setting some of the fields as server-based page fields before the data is actually read.
    - `DefaultVersion As Variant` (optional): The version of Microsoft Excel that the PivotTable was originally created in.
