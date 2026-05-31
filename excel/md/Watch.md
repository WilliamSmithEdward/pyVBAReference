# Watch

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024457-0000-0000-C000-000000000046}  

Represents a range that is tracked when the worksheet is recalculated. The Watch object allows users to verify the accuracy of their models and debug problems that they encounter.

**Remarks:** The Watch object is a member of the Watches collection.

**Example:**

```vba
Sub AddWatch()

 With Application
 .Range("A1").Formula = 1
 .Range("A2").Formula = 2
 .Range("A3").Formula = "=Sum(A1:A2)"
 .Range("A3").Select
 .Watches.Add Source:=ActiveCell
 End With

End Sub
```

## Properties (4)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Source As Variant  (read-only)`  
  Returns a Variant value that represents the unique name that identifies items that have a SourceType property value of xlSourceRange, xlSourceChart, xlSourcePrintArea, xlSourceAutoFilter, xlSourcePivotTable, or xlSourceQuery.

## Methods (1)

- `Delete()`  
  Deletes the object.
