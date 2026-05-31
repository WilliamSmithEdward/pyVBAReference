# PivotFilters

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024484-0000-0000-C000-000000000046}  

The PivotFilters object is a collection of PivotFilter objects.

**Remarks:** The PivotFilters collection contains properties and methods to add new filters, count the number of existing filters in the collection, and reference specific PivotFilter objects.

**Example:**

```vba
ActiveCell.PivotField.PivotFilters.Add FilterType := xlThisWeek
```

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified PivotFilters object. Read-only.
- `_Default As PivotFilter  (read-only)`
- `_NewEnum As IUnknown  (read-only)`
- `Item As PivotFilter  (read-only)`  
  Returns a specific element of the PivotFilters collection object by its position in the collection. Read-only.
- `Count As Long  (read-only)`  
  Returns the number of items in the PivotFilters collection. Read-only.

## Methods (1)

- `Add2(Type As XlPivotFilterType, [DataField As Variant], [Value1 As Variant], [Value2 As Variant], [Order As Variant], [Name As Variant], [Description As Variant], [MemberPropertyField As Variant], [WholeDayFilter As Variant]) As PivotFilter`
