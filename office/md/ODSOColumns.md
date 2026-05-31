# ODSOColumns

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C1532-0000-0000-C000-000000000046}  

A collection of ODSOColumn objects that represent the data fields in a mail merge data source.

**Example:**

```vba
Sub ShowFieldNames()
 Dim appOffice As OfficeDataSourceObject
 Dim intCount As Integer

 Set appOffice = Application.OfficeDataSourceObject
 appOffice.Open bstrConnect:="DRIVER=SQL Server;SERVER=ServerName;" & _
 "UID=user;PWD=;DATABASE=Northwind", bstrTable:="Employees"

 With appOffice.Columns
 For intCount = 1 To .Count
 MsgBox "Column Name: " & .Item(intCount).Name
 Next
 End With
End Sub
```

## Properties (4)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the ODSOColumns object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the ODSOColumns object was created. Read-only.
- `Count As Long  (read-only)`  
  Gets a Long indicating the number of items in the ODSOColumns collection. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the ODSOColumns object. Read-only.

## Methods (1)

- `Item(varIndex As Variant) As Object`  
  Specifies an ODSOColumn object in the ODSOColumns collection.
    - `varIndex As Variant` (required): The index number of the item.
