# XmlMaps

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002447C-0000-0000-C000-000000000046}  

Represents the collection of XmlMap objects that have been added to a workbook.

**Example:**

```vba
Sub AddXmlMap()
 Dim strSchemaLocation As String

 strSchemaLocation = "https://example.microsoft.com/schemas/CustomerData.xsd"
 ActiveWorkbook.XmlMaps.Add strSchemaLocation, "Root"
End Sub
```

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `_Default As XmlMap  (read-only)`
- `Item As XmlMap  (read-only)`  
  Returns a single object from a collection.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.
- `_NewEnum As IUnknown  (read-only)`

## Methods (1)

- `Add(Schema As String, [RootElementName As Variant]) As XmlMap`  
  Adds an XML map to the specified workbook.
    - `Schema As String` (required): The schema to be added as an XML map. The string can be a path to a schema file, or the schema itself. The path can be specified in the Universal Naming Convention (UNC) or Uniform Resource Locator (URL) format.
    - `RootElementName As Variant` (optional): The name of the root element. This argument can be ignored if the schema contains only one root element.
