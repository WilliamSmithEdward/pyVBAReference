# CustomProperties

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024452-0000-0000-C000-000000000046}  

A collection of CustomProperty objects that represents additional information. The information can be used as metadata for XML.

**Remarks:** Use the CustomProperties property of the Worksheet object to return a CustomProperties collection. After a CustomProperties collection is returned, you can add metadata to worksheets and perform additional actions depending on which you choose to work with. To add metadata to a worksheet, use the CustomProperties property with the Add method.

**Example:**

```vba
Sub CheckCustomProperties()

 Dim wksSheet1 As Worksheet

 Set wksSheet1 = Application.ActiveSheet

 ' Add metadata to worksheet.
 wksSheet1.CustomProperties.Add _
 Name:="Market", Value:="Nasdaq"

 ' Display metadata.
 With wksSheet1.CustomProperties.Item(1)
 MsgBox .Name & vbTab & .Value
 End With

End Sub
```

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.
- `_Default As CustomProperty  (read-only)`
- `Item As CustomProperty  (read-only)`  
  Returns a single object from a collection.
- `_NewEnum As IUnknown  (read-only)`

## Methods (1)

- `Add(Name As String, Value As Variant) As CustomProperty`  
  Adds custom property information.
    - `Name As String` (required): The name of the custom property.
    - `Value As Variant` (required): The value of the custom property.
