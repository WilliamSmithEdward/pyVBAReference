# AdditionalData

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {DBC51761-A8ED-11D3-A0DD-00C04F68712B}  

Represents the collection of tables and queries that will be included with the parent table that is exported by the ExportXML method.

**Remarks:** To create an AdditionalData object, use the CreateAdditionalData method of the Application object. To add a table to an existing AdditionalData object, use the Add method.

**Example:**

```vba
Sub ExportCustomerOrderData()
 Dim objOrderInfo As AdditionalData

 Set objOrderInfo = Application.CreateAdditionalData

 ' Add the Orders and Order Details tables to the data to be exported.
 objOrderInfo.Add "Orders"
 objOrderInfo.Add "Order Details"

 ' Export the contents of the Customers table. The Orders and Order
 ' Details tables will be included in the XML file.
 Application.ExportXML ObjectType:=acExportTable, DataSource:="Customers", _
 DataTarget:="Customer Orders.xml", _
 AdditionalData:=objOrderInfo
End Sub
```

## Properties (3)

- `Name As String  (read/write)`  
  Use the Name property to specify or determine the string expression that identifies the name of an object. Read/write String.
- `Item As _AdditionalData  (read-only)`  
  The Item property returns a specific member of a collection either by position or by index. Read-only AdditionalData.
- `Count As Long  (read-only)`  
  Use the Count property to determine the number of items in a specified collection. Read-only Long.

## Methods (1)

- `Add(var As String) As _AdditionalData`  
  Adds an add table or query that will be included when the ExportXML method is called.
    - `var As String` (required): The name of the table or query to add.
