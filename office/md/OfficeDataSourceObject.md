# OfficeDataSourceObject

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C1530-0000-0000-C000-000000000046}  

Represents the mail merge data source in a mail merge operation.

**Remarks:** To work with the OfficeDataSourceObject object, dimension a variable as an OfficeDataSourceObject object. You can then work with the different properties and methods associated with the object. Use the SetSortOrder method to specify how to sort the records in a data source.

**Example:**

```vba
Sub SetDataSortOrder()
 Dim appOffice As OfficeDataSourceObject

 Set appOffice = Application.OfficeDataSourceObject
 appOffice.Open bstrConnect:="DRIVER=SQL Server;SERVER=ServerName;" & _
 "UID=user;PWD=;DATABASE=Northwind", bstrTable:="Employees"

 appOffice.SetSortOrder SortField1:="ZipCode", _
 SortAscending1:=False, SortField2:="LastName", _
 SortField3:="FirstName"
End Sub
```

## Properties (6)

- `ConnectString As String  (read/write)`  
  Gets or sets a String that represents the connection to the specified mail merge data source. Read/write.
- `Table As String  (read/write)`  
  Gets a String that represents the name of the table within the data source file that contains the mail merge records. The returned value may be blank if the table name is unknown or not applicable to the current data source. Read-only.
- `DataSource As String  (read/write)`  
  Gets or sets a String that represents the name of the attached data source. Read/write.
- `Columns As Object  (read-only)`  
  Gets an ODSOColumns object that represents the fields in a data source. Read-only.
- `RowCount As Long  (read-only)`  
  Gets a Long that represents the number of records in the specified data source. Read-only.
- `Filters As Object  (read-only)`  
  Gets the filter status for an OfficeDataSourceObject object. Read-only.

## Methods (4)

- `Move(MsoMoveRow As MsoMoveRow, [RowNbr As Long]) As Long`  
  Moves a record in a return set from an OfficeDataSourceObject object from one position to another.
    - `MsoMoveRow As MsoMoveRow` (required): A constant specifying which row to move.
    - `RowNbr As Long` (optional): The number of the destination row.
- `Open([bstrSrc As String], [bstrConnect As String], [bstrTable As String], [fOpenExclusive As Long], [fNeverPrompt As Long])`  
  Opens a table in an OfficeDataSourceObject object.
    - `bstrSrc As String` (optional): Contains the name of the data source.
    - `bstrConnect As String` (optional): Contains the connection string to the data source.
    - `bstrTable As String` (optional): Specifies which table to open.
    - `fOpenExclusive As Long` (optional): Indicates whether the table should be opened for exclusive access.
    - `fNeverPrompt As Long` (optional): Indicates whether to notify the user if the table cannot be opened.
- `SetSortOrder(SortField1 As String, [SortAscending1 As Boolean], [SortField2 As String], [SortAscending2 As Boolean], [SortField3 As String], [SortAscending3 As Boolean])`  
  Sets the sort order for mail merge data.
    - `SortField1 As String` (required): The first field on which to sort the mail merge data.
    - `SortAscending1 As Boolean` (optional): True (default) to perform an ascending sort on SortField1; False to perform a descending sort.
    - `SortField2 As String` (optional): The second field on which to sort the mail merge data. Default is an empty string.
    - `SortAscending2 As Boolean` (optional): True (default) to perform an ascending sort on SortField2; False to perform a descending sort.
    - `SortField3 As String` (optional): The third field on which to sort the mail merge data. Default is an empty string.
    - `SortAscending3 As Boolean` (optional): True (default) to perform an ascending sort on SortField3; False to perform a descending sort.
- `ApplyFilter()`  
  Applies a filter to a mail merge data source to filter specified records meeting specified criteria.
