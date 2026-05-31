# SortFields

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244AA-0000-0000-C000-000000000046}  

The SortFields collection is a collection of SortField objects. It allows developers to store a sort state on workbooks, lists, and autofilters.

**Remarks:** The object contains properties to add, count, sort, and remove SortField objects.

**Example:**

```vba
ActiveWorksheet.SortFields.Add Key:=Range("A1"), Order:=xlDescending
ActiveWorksheet.SortFields.Add Key:=Range("B1"), Order:=xlDescending
ActiveWorksheet.SortFields.Sort Header:=xlGuess
```

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Item As SortField  (read-only)`  
  Returns a SortField object that represents a collection of items that can be sorted in a workbook. Read-only.
- `Count As Long  (read-only)`  
  Returns the number of objects in the collection. Read-only Long.
- `_Default As SortField  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (3)

- `Add(Key As Range, [SortOn As Variant], [Order As Variant], [CustomOrder As Variant], [DataOption As Variant]) As SortField`  
  Creates a new sort field and returns a SortFields object.
    - `Key As Range` (required): Specifies a key value for the sort.
    - `SortOn As Variant` (optional): An XlSortOn value that specifies which property of a cell to use for the sort.
    - `Order As Variant` (optional): An XlSortOrder value that specifies the sort order.
    - `CustomOrder As Variant` (optional): Specifies if a custom sort order should be used.
    - `DataOption As Variant` (optional): An XlSortDataOption value that specifies how to sort text.
- `Clear()`  
  Clears all the SortFields objects.
- `Add2(Key As Range, [SortOn As Variant], [Order As Variant], [CustomOrder As Variant], [DataOption As Variant], [SubField As Variant]) As SortField`  
  Creates a new sort field and returns a SortFields object that can optionally sort data types with the SubField defined.
    - `Key As Range` (required): Specifies a key value for the sort.
    - `SortOn As Variant` (optional): An XlSortOn value that specifies which property of a cell to use for the sort.
    - `Order As Variant` (optional): An XlSortOrder value that specifies the sort order.
    - `CustomOrder As Variant` (optional): Specifies if a custom sort order should be used.
    - `DataOption As Variant` (optional): An XlSortDataOption value that specifies how to sort text.
    - `SubField As Variant` (optional): Specifies the field to sort on for a data type (such as Population for Geography or Volume for Stocks).
