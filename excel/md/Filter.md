# Filter

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024434-0000-0000-C000-000000000046}  

Represents a filter for a single column.

**Remarks:** The Filter object is a member of the Filters collection. The Filters collection contains all the filters in an autofiltered range.

**Example:**

```vba
Set w = Worksheets("Crew")
If w.AutoFilterMode Then
 filterIsOn = w.AutoFilter.Filters(1).On
End If
```

## Properties (8)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `On As Boolean  (read-only)`  
  True if the specified filter is on. Read-only Boolean.
- `Criteria1 As Variant  (read-only)`  
  Returns the first filtered value for the specified column in a filtered range. Read-only Variant.
- `Criteria2 As Variant  (read-only)`  
  Returns the second filtered value for the specified column in a filtered range. Read-only Variant.
- `Operator As XlAutoFilterOperator  (read/write)`  
  Returns an XlAutoFilterOperator value that represents the operator that associates the two criteria applied by the specified filter.
- `Count As Long  (read-only)`  
  Returns the number of objects in the collection. Read-only Long.
