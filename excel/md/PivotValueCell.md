# PivotValueCell

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244CF-0000-0000-C000-000000000046}  

Provides a way to expose values of cells in the case that actual cells (Range objects) are not available.

**Example:**

```vba
Sub TestEquality()
Dim X As Double
Dim Y As Double

'This code assumes that you have a Standalone PivotChart on one of the worksheets.
X = ThisWorkbook.PivotTables(1).PivotValueCell(1, 1).Value
Y = ThisWorkbook.PivotTables(1).PivotValueCell(1, 2).Value

If X > Y Then
MsgBox "X is greater than Y"
Else
MsgBox "Y is greater than X"
End If
End Sub
```

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Excel application. Read-only.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified PivotValueCell object. Read-only.
- `PivotCell As PivotCell  (read-only)`  
  Returns the PivotCell object that specifies the location of the PivotValueCell. Read-only.
- `Value As Variant  (read-only)`  
  Returns the value at the location. The value is the value after ShowAs and other calculations have been applied. Variant can be Empty, Number, Date, String, or Error value.
- `ServerActions As Actions  (read-only)`  
  Returns a collection of OLAP Action name objects that represent OLAP-defined actions that can be executed. Read-only.

## Methods (1)

- `ShowDetail()`  
  Puts the individual rows of an OLAP data source that contribute to an aggregate onto their own Excel worksheet.
