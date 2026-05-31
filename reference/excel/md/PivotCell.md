# PivotCell

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024458-0000-0000-C000-000000000046}  

Represents a cell in a PivotTable report.

**Remarks:** Use the PivotCell property of the Range collection to return a PivotCell object. After a PivotCell object is returned, you can use the ColumnItems or RowItems property to determine the PivotItems collection that corresponds to the items on the column or row axis that represents the selected number.

**Example:**

```vba
Sub CheckPivotCellType()

 On Error GoTo Not_In_PivotTable

 ' Determine if cell A5 is a data item in the PivotTable.
 If Application.Range("A5").PivotCell.PivotCellType = xlPivotCellValue Then
 MsgBox "The PivotCell at A5 is a data item."
 Else
 MsgBox "The PivotCell at A5 is not a data item."
 End If
 Exit Sub

Not_In_PivotTable:
 MsgBox "The chosen cell is not in a PivotTable."

End Sub
```

## Properties (18)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `PivotCellType As XlPivotCellType  (read-only)`  
  Returns one of the XlPivotCellType constants that identifies the PivotTable entity that the cell corresponds to. Read-only.
- `PivotTable As PivotTable  (read-only)`  
  Returns a PivotTable object that represents the PivotTable report associated with the PivotCell.
- `DataField As PivotField  (read-only)`  
  Returns a PivotField object that corresponds to the selected data field.
- `PivotField As PivotField  (read-only)`  
  Returns a PivotField object that represents the PivotTable field containing the upper-left corner of the specified range.
- `PivotItem As PivotItem  (read-only)`  
  Returns a PivotItem object that represents the PivotTable item containing the upper-left corner of the specified range.
- `RowItems As PivotItemList  (read-only)`  
  Returns a PivotItemList collection that corresponds to the items on the category axis that represent the selected cell.
- `ColumnItems As PivotItemList  (read-only)`  
  Returns a PivotItemList collection that corresponds to the items on the column axis that represent the selected range.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the range that the specified PivotCell applies to.
- `CustomSubtotalFunction As XlConsolidationFunction  (read-only)`  
  Returns the custom subtotal function field setting of a PivotCell object. Read-only XlConsolidationFunction.
- `PivotRowLine As PivotLine  (read-only)`  
  Returns the PivotLine object on a row for a specific PivotCell object. Read-only PivotLine.
- `PivotColumnLine As PivotLine  (read-only)`  
  Returns the PivotLine object on a column for a specific PivotCell object. Read-only PivotLine.
- `DataSourceValue As Variant  (read-only)`  
  Returns the value last retrieved from the data source for edited cells in a PivotTable report. Read-only.
- `CellChanged As XlCellChangedState  (read-only)`  
  Returns whether a PivotTable value cell has been edited or recalculated since the PivotTable report was created or the last commit operation was performed. Read-only.
- `MDX As String  (read-only)`  
  Returns a tuple that provides the full MDX coordinates of the specified value cell in a PivotTable with an OLAP data source. Read-only.
- `ServerActions As Actions  (read-only)`  
  Represents a collection of _actions_ consisting of OLAP-defined actions that can be executed. The actions are specific to PivotTables existing at a worksheet-level. Read-only.

## Methods (2)

- `AllocateChange()`  
  Performs a writeback operation on the specified cell in a PivotTable report based on an OLAP data source.
- `DiscardChange()`  
  Discards changes to the specified cell in a PivotTable report.
