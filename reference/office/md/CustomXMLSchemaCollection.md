# CustomXMLSchemaCollection

**Type:** Class  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000CDB0D-0000-0000-C000-000000000046}  

Represents a collection of CustomXMLSchema objects attached to a data stream.

**Example:**

```vba
Dim SchemaCollection As CustomXMLSchemaCollection

SchemaCollection.Add "https://tempuri.org/XMLSchema.xsd"
```

## Properties (7)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the CustomXMLSchemaCollection object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the CustomXMLSchemaCollection object was created. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the CustomXMLSchemaCollection object. Read-only.
- `Count As Long  (read-only)`  
  Gets a Long indicating the number of items in the CustomXMLSchemaCollection collection. Read-only.
- `Item As CustomXMLSchema  (read-only)`  
  Gets a CustomXMLSchema object from the CustomXMLSchemaCollection collection. Read-only.
- `NamespaceURI As String  (read-only)`  
  Gets the unique address identifier for the namespace of the CustomXMLSchemaCollection object. Read-only.
- `_NewEnum As IUnknown  (read-only)`

## Methods (3)

- `Add([NamespaceURI As String], [Alias As String], [FileName As String], [InstallForAllUsers As Boolean]) As CustomXMLSchema`  
  Allows you to add one or more schemas to a schema collection that can then be added to a stream in the data store and to the Schema Library.
    - `NamespaceURI As String` (optional): Contains the namespace of the schema to be added to the collection. If the schema already exists in the Schema Library, the method will retrieve it from there.
    - `Alias As String` (optional): Contains the alias of the schema to be added to the collection. If the alias already exists in the Schema Library, the method can find it using this argument.
    - `FileName As String` (optional): Contains the location of the schema on a disk. If this parameter is specified, the schema is added to the collection and to the Schema Library.
    - `InstallForAllUsers As Boolean` (optional): Specifies whether, in the case where the method is adding the schema to the Schema Library, the Schema Library keys should be written to the registry (HKey_Local_Machine for all users or HKey_Current_User for just the current user). The parameter defaults to False and writes to HKey_Current_User.
- `AddCollection(SchemaCollection As CustomXMLSchemaCollection)`  
  Adds an existing schema collection to the current schema collection.
    - `SchemaCollection As CustomXMLSchemaCollection` (required): Represents a collection of schemas to be imported into the current schema collection.
- `Validate() As Boolean`  
  Specifies whether the schemas in a schema collection are valid (conforms to the syntactic rules of XML and the rules for a specified vocabulary; a standard for structuring XML).
