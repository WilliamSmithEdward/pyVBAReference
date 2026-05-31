# PivotCaches

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002441D-0000-0000-C000-000000000046}  

Represents the collection of memory caches from the PivotTable reports in a workbook.

**Remarks:** Each memory cache is represented by a PivotCache object.

**Example:**

```vba
For Each pc In ActiveWorkbook.PivotCaches
 pc.RefreshOnFileOpen = True
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
- `_Default As PivotCache  (read-only)`

## Methods (3)

- `Item(Index As Variant) As PivotCache`  
  Returns a single object from a collection.
    - `Index As Variant` (required): The name or index number for the object.
- `_NewEnum() As IUnknown`
- `Create(SourceType As XlPivotTableSourceType, [SourceData As Variant], [Version As Variant]) As PivotCache`  
  Creates a new PivotCache.
    - `SourceType As XlPivotTableSourceType` (required): _SourceType_ can be one of these XlPivotTableSourceType constants: xlConsolidation, xlDatabase, or xlExternal.
    - `SourceData As Variant` (optional): The data for the new PivotTable cache.
    - `Version As Variant` (optional): Version of the PivotTable. _Version_ can be one of the XlPivotTableVersionList constants.
