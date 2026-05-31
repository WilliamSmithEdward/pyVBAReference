# CustomXMLSchema

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000CDB01-0000-0000-C000-000000000046}  

Represents a schema in a CustomXMLSchemaCollection collection.

**Example:**

```vba
Dim SchemaCollection As CustomXMLSchemaCollection

SchemaCollection.Add "https://tempuri.org/XMLSchema.xsd"
```

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the CustomXMLSchema object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the CustomXMLSchema object was created. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the CustomXMLSchema object. Read-only.
- `Location As String  (read-only)`  
  Gets a String that represents the location of a schema on a computer. Read-only.
- `NamespaceURI As String  (read-only)`  
  Gets the unique address identifier for the namespace of the CustomXMLSchema object. Read-only.

## Methods (2)

- `Delete()`  
  Deletes the specified schema from the CustomXMLSchema collection.
- `Reload()`  
  Reloads a schema from a file.
