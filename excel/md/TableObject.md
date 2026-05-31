# TableObject

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244CE-0000-0000-C000-000000000046}  

Represents a worksheet table built from data returned from a PowerPivot model.

**Example:**

```vba
Sub CreateTable()
Dim objWBConnection As WorkbookConnection
Dim objWorksheet As Worksheet
Dim objTable As TableObject   'This is the new Table object

Set objWorksheet = ActiveWorkbook.Worksheets("Sheet1")

'Create a WorkbookConnection to the external data source first.
Set objWBConnection = ActiveWorkbook.Connections.Add2( _
        "Cubes3 AdventureWorksDW DimEmployee1", "", Array( _
        "OLEDB;Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=AdventureWorksDW;Data Source=MyServer;Use " _
        , _
        "Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=MYWORKSTATION;Use Encryption for Data=False;Tag with co" _
        , "lumn collation when possible=False"), Array( _
        """AdventureWorksDW"".""dbo"".""DimEmployee"""), 3, True)

'Create a new table connected to the model.
Set objTable = objWorksheet.ListObjects.Add(SourceType:=xlSrcModel, Source:=objWBConnection, Destination:=Range("$A$1")).TableObject

objTable.Refresh

End Sub
```

## Properties (15)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Excel application. Read-only.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified TableObject object. Read-only.
- `RowNumbers As Boolean  (read/write)`  
  Specifies if row numbers are added as the first column of the specified query table. Read/write Boolean.
- `FetchedRowOverflow As Boolean  (read-only)`  
  Specifies if the number of rows returned by the last use of the Refresh method is greater than the number of rows available on the worksheet. Read-only Boolean.
- `RefreshStyle As XlCellInsertionMode  (read/write)`  
  Returns or sets the way rows on the specified worksheet are added or deleted to accommodate the number of rows in a record set returned by a query. Read/write XlCellInsertionMode enumeration.
- `EnableRefresh As Boolean  (read/write)`  
  Specifies if the query table can be refreshed by the user. Read/write Boolean.
- `Destination As Range  (read-only)`  
  Returns the cell in the upper-left corner of the query table destination range (the range where the resulting query table will be placed). The destination range must be on the worksheet that contains the TableObject object. Read-only Range.
- `ResultRange As Range  (read-only)`  
  Returns a Range object that represents the area of the worksheet occupied by the specified query table. Read-only.
- `EnableEditing As Boolean  (read/write)`  
  True if the user can edit the specified query table. False if the user can only refresh the query table. Read/write Boolean.
- `PreserveColumnInfo As Boolean  (read/write)`  
  Specifies if column sorting, filtering, and layout information is preserved whenever a query table is refreshed. The default value is False. Read/write Boolean.
- `PreserveFormatting As Boolean  (read/write)`  
  True if any formatting common to the first five rows of data are applied to new rows of data in the query table. Unused cells aren't formatted. The property is False if the last AutoFormat applied to the query table is applied to new rows of data. The default value is True. Read/write Boolean.
- `AdjustColumnWidth As Boolean  (read/write)`  
  Specifies if the column widths are automatically adjusted for the best fit each time you refresh the specified query table. The default value is True. Read/write Boolean.
- `ListObject As ListObject  (read-only)`  
  Returns a ListObject object for the TableObject object. Read-only.
- `WorkbookConnection As WorkbookConnection  (read-only)`  
  Returns the WorkbookConnection object used by the TableObject for connecting to the model.

## Methods (2)

- `Delete()`  
  Deletes the TableObject object.
- `Refresh() As Boolean`  
  This method updates the TableObject object.
