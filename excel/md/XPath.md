# XPath

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002447E-0000-0000-C000-000000000046}  

Represents an XPath that has been mapped to a Range or ListColumn object.

**Remarks:** Use the SetValue method to map an XPath to a range or list column. The SetValue method is also used to change the properties of an existing XPath. Use the Clear method to remove an XPath that has been mapped to a range or list column.

**Example:**

```vba
Sub CreateXMLList()
 Dim mapContact As XmlMap
 Dim strXPath As String
 Dim lstContacts As ListObject
 Dim lcNewCol As ListColumn

 ' Specify the schema map to use.
 Set mapContact = ActiveWorkbook.XmlMaps("Contacts")

 ' Create a new list.
 Set lstContacts = ActiveSheet.ListObjects.Add

 ' Specify the first element to map.
 strXPath = "/Root/Person/FirstName"
 ' Map the element.
 lstContacts.ListColumns(1).XPath.SetValue mapContact, strXPath

 ' Specify the element to map.
 strXPath = "/Root/Person/LastName"
 ' Add a column to the list.
 Set lcNewCol = lstContacts.ListColumns.Add
 ' Map the element.
 lcNewCol.XPath.SetValue mapContact, strXPath

 strXPath = "/Root/Person/Address/Zip"
 Set lcNewCol = lstContacts.ListColumns.Add
 lcNewCol.XPath.SetValue mapContact, strXPath
End Sub
```

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `_Default As String  (read-only)`
- `Value As String  (read-only)`  
  Returns a String that represents the XPath for the specified object.
- `Map As XmlMap  (read-only)`  
  Returns an XmlMap object that represents the schema map that contains the specified XPath object. Read-only.
- `Repeating As Boolean  (read-only)`  
  Returns True if the specified XPath object is mapped to an XML list; returns False if the XPath object is mapped to a single cell. Read-only Boolean.

## Methods (2)

- `SetValue(Map As XmlMap, XPath As String, [SelectionNamespace As Variant], [Repeating As Variant])`  
  Maps the specified XPath object to a ListColumn object or Range collection. If the XPath object has previously been mapped to the ListColumn object or Range collection, the SetValue method sets the properties of the XPath object.
    - `Map As XmlMap` (required): The map info that the mapped range will be associated with.
    - `XPath As String` (required): A valid XPath expression that tells Excel what XML data should appear in this mapped range. The XPath string can also contain valid filters, in which case, only a subset of the data that the XPath points to will ever appear in this mapped range.
    - `SelectionNamespace As Variant` (optional): Specifies any namespace prefixes used in the _XPath_ argument. This argument can be omitted if the XPath object doesn't contain any prefixes, or if the XPath object uses the Microsoft Excel internal prefixes.
    - `Repeating As Variant` (optional): Specifies whether the XPath object is to be bound to a column in an XML list or mapped to a single cell. Set to True to bind the XPath object to a column in an XML list. False forces a non-repeating cell to be created. If the range is greater than a single cell and False is specified, a run-time error occurs.
- `Clear()`  
  Clears all XPath schema information for the mapped range.
