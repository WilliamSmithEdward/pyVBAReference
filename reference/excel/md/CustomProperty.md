# CustomProperty

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024453-0000-0000-C000-000000000046}  

Represents identifier information, which can be used as metadata for XML.

**Remarks:** Use the Add method or the Item property of the CustomProperties collection to return a CustomProperty object. After a CustomProperty object is returned, you can add metadata to worksheets by using the CustomProperties property of the Worksheet object with the Add method.

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

## Properties (6)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Name As String  (read/write)`  
  Returns or sets a String value representing the name of the object.
- `Value As Variant  (read/write)`  
  Synonym for the Borders.LineStyle property.
- `_Default As Variant  (read-only)`

## Methods (1)

- `Delete()`  
  Deletes the object.
