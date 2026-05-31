# ODSOColumn

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C1531-0000-0000-C000-000000000046}  

Represents a field in a data source. The ODSOColumn object is a member of the ODSOColumns collection.

**Remarks:** The ODSOColumns collection includes all the data fields in a mail merge data source (for example, Name, Address, and City). You cannot add fields to the ODSOColumns collection. All data fields in a data source are automatically included in the ODSOColumns collection. Use Columns(_index_), where _index_ is the data field name or index number, to return a single ODSOColumn object. The index number represents the position of the data field in the mail merge data source.

**Example:**

```vba
Sub GetDataFromSource()

 Dim appOffice As OfficeDataSourceObject

 Set appOffice = Application.OfficeDataSourceObject
 appOffice.Open bstrConnect:="DRIVER=SQL Server;SERVER=ServerName;" & _
 "UID=user;PWD=;DATABASE=Northwind", bstrTable:="Employees"

 With appOffice.Columns
 MsgBox "Field Name: " & .Item(1).Name & vbLf & _
 "Value: " & .Item(1).Value
 End With
End Sub
```

## Properties (6)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the ODSOColumn object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the ODSOColumn object was created. Read-only.
- `Index As Long  (read-only)`  
  Gets a Long representing the index number for an ODSOColumn object in the collection. Read-only.
- `Name As String  (read-only)`  
  Gets the name of a data field in a mail merge data source. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the ODSOColumn object. Read-only.
- `Value As String  (read-only)`  
  Gets the value of a data field in a mail merge data source. Read-only.
