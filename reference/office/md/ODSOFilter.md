# ODSOFilter

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C1533-0000-0000-C000-000000000046}  

Represents a filter to be applied to an attached mail merge data source. The ODSOFilter object is a member of the ODSOFilters object.

**Remarks:** Each filter is a line in a query string. Use the Column, CompareTo, Comparison, and Conjunction properties to return or set the data source query criterion.

**Example:**

```vba
Sub SetQueryCriterion()
 Dim appOffice As Office.OfficeDataSourceObject
 Dim intItem As Integer

 Set appOffice = Application.OfficeDataSourceObject
 appOffice.Open bstrConnect:="DRIVER=SQL Server;SERVER=ServerName;" & _
 "UID=user;PWD=;DATABASE=Northwind", bstrTable:="Employees"

 With appOffice.Filters
 For intItem = 1 To .Count
 With .Item(intItem)
 If .Column = "Region" Then
 .Comparison = msoFilterComparisonNotEqual
 .CompareTo = "WA"
 If .Conjunction = "Or" Then .Conjunction = "And"
 End If
 End With
 Next intItem
 End With
End Sub
```

## Properties (8)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the ODSOFilter object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the ODSOFilter object was created. Read-only.
- `Index As Long  (read-only)`  
  Gets a Long representing the index number for an ODSOFilter object in the collection. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the ODSOFilter object. Read-only.
- `Column As String  (read/write)`  
  Gets or sets a String that represents the name of the field in the mail merge data source to use in the filter. Read/write.
- `Comparison As MsoFilterComparison  (read/write)`  
  Gets or sets an MsoFilterComparison constant that represents how to compare the Column and CompareTo properties. Read/write.
- `CompareTo As String  (read/write)`  
  Gets or sets a String that represents the text to compare in the query filter criterion. Read/write.
- `Conjunction As MsoFilterConjunction  (read/write)`  
  Gets or sets an MsoFilterConjunction constant that represents how a filter criterion relates to other filter criteria in the ODSOFilters object. Read/write.
