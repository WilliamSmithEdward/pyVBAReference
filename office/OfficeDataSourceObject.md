# OfficeDataSourceObject

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C1530-0000-0000-C000-000000000046}  

## Properties (6)

- `ConnectString As String  (read/write)`
- `Table As String  (read/write)`
- `DataSource As String  (read/write)`
- `Columns As Object  (read-only)`
- `RowCount As Long  (read-only)`
- `Filters As Object  (read-only)`

## Methods (4)

- `Move(MsoMoveRow As MsoMoveRow, [RowNbr As Long]) As Long`
- `Open([bstrSrc As String], [bstrConnect As String], [bstrTable As String], [fOpenExclusive As Long], [fNeverPrompt As Long])`
- `SetSortOrder(SortField1 As String, [SortAscending1 As Boolean], [SortField2 As String], [SortAscending2 As Boolean], [SortField3 As String], [SortAscending3 As Boolean])`
- `ApplyFilter()`
