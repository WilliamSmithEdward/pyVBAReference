# Protection

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024467-0000-0000-C000-000000000046}  

Represents the various types of protection options available for a worksheet.

**Remarks:** Use the Protection property of the Worksheet object to return a Protection object. After a Protection object is returned, you can use the Protection properties to set or return protection options.

**Example:**

```vba
Sub SetProtection()

 Range("A1").Formula = "1"
 Range("B1").Formula = "3"
 Range("C1").Formula = "4"
 ActiveSheet.Protect

 ' Check the protection setting of the worksheet and act accordingly.
 If ActiveSheet.Protection.AllowInsertingColumns = False Then
 ActiveSheet.Protect AllowInsertingColumns:=True
 MsgBox "Insert a column between 1 and 3"
 Else
 MsgBox "Insert a column between 1 and 3"
 End If

End Sub
```

## Properties (12)

- `AllowFormattingCells As Boolean  (read-only)`  
  Returns True if the formatting of cells is allowed on a protected worksheet. Read-only Boolean.
- `AllowFormattingColumns As Boolean  (read-only)`  
  Returns True if the formatting of columns is allowed on a protected worksheet. Read-only Boolean.
- `AllowFormattingRows As Boolean  (read-only)`  
  Returns True if the formatting of rows is allowed on a protected worksheet. Read-only Boolean.
- `AllowInsertingColumns As Boolean  (read-only)`  
  Returns True if the insertion of columns is allowed on a protected worksheet. Read-only Boolean.
- `AllowInsertingRows As Boolean  (read-only)`  
  Returns True if the insertion of rows is allowed on a protected worksheet. Read-only Boolean.
- `AllowInsertingHyperlinks As Boolean  (read-only)`  
  Returns True if the insertion of hyperlinks is allowed on a protected worksheet. Read-only Boolean.
- `AllowDeletingColumns As Boolean  (read-only)`  
  Returns True if the deletion of columns is allowed on a protected worksheet. Read-only Boolean.
- `AllowDeletingRows As Boolean  (read-only)`  
  Returns True if the deletion of rows is allowed on a protected worksheet. Read-only Boolean.
- `AllowSorting As Boolean  (read-only)`  
  Returns True if the sorting option is allowed on a protected worksheet. Read-only Boolean.
- `AllowFiltering As Boolean  (read-only)`  
  Returns True if the user is allowed to make use of an AutoFilter that was created before the sheet was protected. Read-only Boolean.
- `AllowUsingPivotTables As Boolean  (read-only)`  
  Returns True if the user is allowed to manipulate PivotTables on a protected worksheet. Read-only Boolean.
- `AllowEditRanges As AllowEditRanges  (read-only)`  
  Returns an AllowEditRanges object.
