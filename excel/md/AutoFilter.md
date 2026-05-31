# AutoFilter

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024432-0000-0000-C000-000000000046}  

Represents autofiltering for the specified worksheet.

**Example:**

```vba
Dim w As Worksheet
Dim filterArray()
Dim currentFiltRange As String

Sub ChangeFilters()

Set w = Worksheets("Crew")
With w.AutoFilter
 currentFiltRange = .Range.Address
 With .Filters
 ReDim filterArray(1 To .Count, 1 To 3)
 For f = 1 To .Count
 With .Item(f)
 If .On Then
 filterArray(f, 1) = .Criteria1
 If .Operator Then
 filterArray(f, 2) = .Operator
 filterArray(f, 3) = .Criteria2
 End If
 End If
 End With
 Next
 End With
End With

w.AutoFilterMode = False
w.Range("A1").AutoFilter field:=1, Criteria1:="S"

End Sub
```

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the range to which the specified AutoFilter applies.
- `Filters As Filters  (read-only)`  
  Returns a Filters collection that represents all the filters in an autofiltered range. Read-only.
- `FilterMode As Boolean  (read-only)`  
  Returns True if the worksheet is in the AutoFilter filter mode. Read-only Boolean.
- `Sort As Sort  (read-only)`  
  Gets the sort column or columns, and sort order for the AutoFilter collection.

## Methods (2)

- `ApplyFilter()`  
  Applies the specified AutoFilter object.
- `ShowAllData()`  
  Displays all the data returned by the AutoFilter object.
