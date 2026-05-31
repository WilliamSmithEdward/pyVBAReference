# ListColumn

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024473-0000-0000-C000-000000000046}  

Represents a column in a table.

**Remarks:** The ListColumn object is a member of the ListColumns collection. The ListColumns collection contains all the columns in a table. Use the ListColumns property of the ListObject object to return a ListColumns collection.

**Example:**

```vba
Sub AddListColumn()
 Dim wrksht As Worksheet
 Dim objListCol As ListColumn

 Set wrksht = ActiveWorkbook.Worksheets("Sheet1")
 Set objListCol = wrksht.ListObjects(1).ListColumns.Add
End Sub
```

## Properties (11)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `_Default As String  (read-only)`
- `Index As Long  (read-only)`  
  Returns a Long value that represents the index number of the ListColumn object within the ListColumns collection.
- `Name As String  (read/write)`  
  Returns or sets a String value that represents the name of the list column.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the range to which the specified list object applies.
- `TotalsCalculation As XlTotalsCalculation  (read/write)`  
  Determines the type of calculation in the Totals row of the list column based on the value of the XlTotalsCalculation enumeration. Read/write.
- `XPath As XPath  (read-only)`  
  Returns an XPath object that represents the XPath of the element mapped to the specified Range object. The context of the range determines whether the action succeeds or returns an empty object. Read-only.
- `DataBodyRange As Range  (read-only)`  
  Returns a Range object that is the size of the data portion of a column. Read-only.
- `Total As Range  (read-only)`  
  Returns the Total row for a ListColumn object. Read-only.

## Methods (1)

- `Delete()`  
  Deletes the column of data in the list.
