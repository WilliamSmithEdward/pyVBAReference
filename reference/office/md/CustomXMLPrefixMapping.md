# CustomXMLPrefixMapping

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000CDB10-0000-0000-C000-000000000046}  

Represents a namespace prefix.

**Example:**

```vba
Dim objNamespace As CustomXMLPrefixMapping

objNamespace = CustomXMLPrefixMappings.AddNamespace("xs", "urn:invoice:namespace")
```

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the CustomXMLPrefixMapping object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the CustomXMLPrefixMapping object was created. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object of the CustomXMLPrefixMapping object. Read-only.
- `Prefix As String  (read-only)`  
  Gets the prefix for a CustomXMLPrefixMapping object. Read-only.
- `NamespaceURI As String  (read-only)`  
  Gets the unique address identifier for the namespace of the CustomXMLPrefixMapping object. Read-only.
