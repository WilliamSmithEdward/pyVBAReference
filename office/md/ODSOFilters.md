# ODSOFilters

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C1534-0000-0000-C000-000000000046}  

Represents all the filters to apply to the data source attached to the mail merge publication. The ODSOFilters object is composed of ODSOFilter objects.

**Remarks:** Use the Add method of the ODSOFilters object to add a new filter criterion to the query.

**Example:**

```vba
Sub SetQueryCriterion()
 Dim appOffice As OfficeDataSourceObject

 Set appOffice = Application.OfficeDataSourceObject
 appOffice.Open bstrConnect:="DRIVER=SQL Server;SERVER=ServerName;" & _
 "UID=user;PWD=;DATABASE=Northwind", bstrTable:="Employees"

 With appOffice.Filters
 .Add Column:="Region", _
 Comparison:=msoFilterComparisonIsBlank, _
 Conjunction:=msoFilterConjunctionAnd
 .ApplyFilter
 End With
End Sub
```

## Properties (4)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the ODSOFilters object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the ODSOFilters object was created. Read-only.
- `Count As Long  (read-only)`  
  Gets a Long indicating the number of items in the ODSOFilters collection. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the ODSOFilters object. Read-only.

## Methods (3)

- `Item(Index As Long) As Object`  
  Represents an ODSOFilter object in the ODSOFilters collection.
    - `Index As Long` (required): The number of the item.
- `Add(Column As String, Comparison As MsoFilterComparison, Conjunction As MsoFilterConjunction, [bstrCompareTo As String], [DeferUpdate As Boolean])`  
  Adds a new filter to the ODSOFilters collection.
    - `Column As String` (required): The name of the table in the data source.
    - `Comparison As MsoFilterComparison` (required): How the data in the table is filtered.
    - `Conjunction As MsoFilterConjunction` (required): Determines how this filter relates to other filters in the ODSOFilters object.
    - `bstrCompareTo As String` (optional): If the _Comparison_ argument is something other than msoFilterComparisonIsBlank or msoFilterComparisonIsNotBlank, _bstrCompareTo_ is a string to which the data in the table is compared.
    - `DeferUpdate As Boolean` (optional): Specifies whether to delay updating the filter. Default is False.
- `Delete(Index As Long, [DeferUpdate As Boolean])`  
  Deletes a filter object from the ODSOFilters collection.
    - `Index As Long` (required): The number of the filter to delete.
    - `DeferUpdate As Boolean` (optional): Specifies whether to delay updating the filter. Default is False.
