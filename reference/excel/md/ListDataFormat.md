# ListDataFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002447D-0000-0000-C000-000000000046}  

The ListDataFormat object holds all the data type properties of the ListColumn object. These properties are read-only.

**Remarks:** Use the ListDataFormat property of the ListColumn object to return a ListDataFormat object. The default property of the ListDataFormat object is the Type property, which indicates the data type of the list column. This allows the user to write code without specifying the Type property. <!--can't find a ListDataFormat property to link to-->

**Example:**

```vba
Dim objListObject As ListObject
Dim objDataRange As Range
Dim strListGUID as String
Dim strServerName as String

strServerName = "https://<servername>/_vti_bin"
strListGUID = "{<listguid>}"

Set objListObject = Sheet1.ListObjects.Add(xlSrcExternal, _
 Array(strServerName, strListGUID), True, xlYes, Range("A1"))

With objListObject.ListColumns(2)
 Set objDataRange = .Range.Offset(1, 0).Resize(.Range.Rows.Count - 2, 1)
 If .ListDataFormat.Type = xlListDataTypeText And .ListDataFormat.Required Then
 objDataRange.Value = "Hello World"
 End If
End With
```

## Properties (16)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `_Default As XlListDataType  (read-only)`
- `Choices As Variant  (read-only)`  
  Returns an Array of String values that contains the choices offered to the user by the ListLookUp, ChoiceMulti, and Choice data types of the DefaultValue property. Read-only Variant.
- `DecimalPlaces As Long  (read-only)`  
  Returns a Long value that represents the number of decimal places to show for the numbers in the ListColumn object. Read-only Long.
- `DefaultValue As Variant  (read-only)`  
  Returns Variant representing the default data type value for a new row in a column. The Nothing object is returned when the schema does not specify a default value. Read-only Variant.
- `IsPercent As Boolean  (read-only)`  
  Returns a Boolean value. Returns True only if the number data for the ListColumn object will be shown in percentage formatting. Read-only Boolean.
- `lcid As Long  (read-only)`  
  Returns a Long value that represents the LCID for the ListColumn object that is specified in the schema definition. Read-only Long.
- `MaxCharacters As Long  (read-only)`  
  Returns a Long containing the maximum number of characters allowed in the ListColumn object if the Type property is set to xlListDataTypeText or xlListDataTypeMultiLineText (XlListDataType enumeration). Read-only Long.
- `MaxNumber As Variant  (read-only)`  
  Returns a Variant containing the maximum value allowed in this field in the list column. Read-only Variant.
- `MinNumber As Variant  (read-only)`  
  Returns a Variant containing the minimum value allowed in this field in the list column. This can be a negative floating point number. Read-only Variant.
- `Required As Boolean  (read-only)`  
  Returns a Boolean value indicating whether the schema definition of a column requires data before the row is committed. Read-only Boolean.
- `Type As XlListDataType  (read-only)`  
  Returns an XlListDataType value that represents the data type of the list column.
- `ReadOnly As Boolean  (read-only)`  
  Returns True if the object has been opened as read-only. Read-only Boolean.
- `AllowFillIn As Boolean  (read-only)`  
  Returns a Boolean value indicating whether users can provide their own data for cells in a column (rather than being restricted to a list of values) for those columns that supply a list of values. Returns False for lists that are not linked to a SharePoint site. Also returns False if the column is not specified as choice or multi-choice. Read-only Boolean.
