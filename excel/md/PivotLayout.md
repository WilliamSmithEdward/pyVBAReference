# PivotLayout

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002444A-0000-0000-C000-000000000046}  

Represents the placement of fields in a PivotChart report.

**Example:**

```vba
Sub ListFieldNames

 Dim objNewSheet As Worksheet
 Dim intRow As Integer
 Dim objPF As PivotField

 Set objNewSheet = Worksheets.Add

 intRow = 1

 For Each objPF In _
 Charts("Chart1").PivotLayout.PivotFields

 objNewSheet.Cells(intRow, 1).Value = objPF.Caption

 intRow = intRow + 1

 Next objPF

End Sub
```

## Properties (4)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `PivotTable As PivotTable  (read-only)`  
  Returns a PivotTable object that represents the PivotTable report associated with the PivotChart report.
