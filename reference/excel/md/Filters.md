# Filters

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024433-0000-0000-C000-000000000046}  

A collection of Filter objects that represents all the filters in an autofiltered range.

**Example:**

```vba
Dim f As Filter
Dim w As Worksheet
Const ns As String = "Not set"

Set w = Worksheets("Crew")
Set w2 = Worksheets("FilterData")
rw = 1
For Each f In w.AutoFilter.Filters
 If f.On Then
 c1 = Right(f.Criteria1, Len(f.Criteria1) - 1)
 If f.Operator Then
 op = f.Operator
 c2 = Right(f.Criteria2, Len(f.Criteria2) - 1)
 Else
 op = ns
 c2 = ns
 End If
 Else
 c1 = ns
 op = ns
 c2 = ns
 End If
 w2.Cells(rw, 1) = c1
 w2.Cells(rw, 2) = op
 w2.Cells(rw, 3) = c2
 rw = rw + 1
Next
```

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.
- `_Default As Filter  (read-only)`
- `Item As Filter  (read-only)`  
  Returns a single object from a collection.
- `_NewEnum As IUnknown  (read-only)`
