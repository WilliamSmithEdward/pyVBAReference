# PivotFilter

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024483-0000-0000-C000-000000000046}  

A PivotFilter is applied to a PivotField object.

**Remarks:** Developers have the option of naming filters for reference because the index is not reliable. The DataField property specifies the PivotField to base a value filter on.

## Properties (16)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified PivotFilter object. Read-only.
- `Order As Long  (read/write)`  
  Specifies the evaluation order of the filter among all Value filters applied to the entire PivotTable. Read/write Integer.
- `FilterType As XlPivotFilterType  (read-only)`  
  Specifies the type of filter to be applied. Read-only XlPivotFilterType.
- `Name As String  (read-only)`  
  This property provides the option of naming filters for reference. You cannot rely on the index value for accurate reference because this value can change.
- `Description As String  (read-only)`  
  Provides an optional description for the PivotFilter object. Read-only String.
- `Active As Boolean  (read-only)`  
  Returns whether the specified PivotFilter is active. Read-only Boolean.
- `PivotField As PivotField  (read-only)`  
  Specifies the PivotField to which the filter is applied. Read-only.
- `DataField As PivotField  (read-only)`  
  This property is applicable only to non-OLAP PivotTables and provides the Value field (PivotField in the Values area) being filtered by for a value filter. Read/write PivotField.
- `DataCubeField As CubeField  (read-only)`  
  This property is applicable only to OLAP PivotTables and provides the Value field (PivotField in the Values area) being filtered by for a value filter. Read/write CubeField.
- `Value1 As Variant  (read-only)`  
  This property is a user-supplied parameter to define a filter for a PivotField. Read/write Variant.
- `Value2 As Variant  (read-only)`  
  This property is a user-supplied parameter to define a filter for a PivotField. Read/write Variant.
- `MemberPropertyField As PivotField  (read-only)`  
  This property specifies the member property PivotField on which the label filter is based. Read/write PivotField.
- `IsMemberPropertyFilter As Boolean  (read-only)`  
  Specifies whether the label filter is based on the PivotItem captions of a member property of the field or on the PivotItem captions of the PivotField itself. Read-only Boolean.
- `WholeDayFilter As Boolean  (read/write)`  
  Sets or gets the filtering semantics for date filters. Read/write Boolean.

## Methods (1)

- `Delete()`  
  Deletes the filter and removes it from the filter collections of the PivotField and the PivotTable.
