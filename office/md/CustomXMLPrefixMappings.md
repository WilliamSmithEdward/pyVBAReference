# CustomXMLPrefixMappings

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000CDB00-0000-0000-C000-000000000046}  

Represents a collection of CustomXMLPrefixMapping objects.

**Example:**

```vba
Dim objNamespace As CustomXMLPrefixMapping

objNamespace = CustomXMLPrefixMappings.AddNamespace("xs", "urn:invoice:namespace")
```

## Properties (6)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the CustomXMLPrefixMappings object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the CustomXMLPrefixMappings object was created. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the CustomXMLPrefixMappings object. Read-only.
- `Count As Long  (read-only)`  
  Gets a Long indicating the number of items in the CustomXMLPrefixMappings collection. Read-only.
- `Item As CustomXMLPrefixMapping  (read-only)`  
  Gets a CustomXMLPrefixMapping object from the CustomXMLPrefixMappings collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`

## Methods (3)

- `AddNamespace(Prefix As String, NamespaceURI As String)`  
  Allows you to add a custom namespace/prefix mapping to use when querying an item.
    - `Prefix As String` (required): Contains the prefix to add to the prefix mapping list.
    - `NamespaceURI As String` (required): Contains the namespace to assign to the newly added prefix.
- `LookupNamespace(Prefix As String) As String`  
  Allows you to get the namespace corresponding to the specified prefix.
    - `Prefix As String` (required): Contains a prefix in the prefix mapping list.
- `LookupPrefix(NamespaceURI As String) As String`  
  Allows you to get a prefix corresponding to the specified namespace.
    - `NamespaceURI As String` (required): Contains the namespace URI.
