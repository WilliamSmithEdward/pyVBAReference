# XMLSchemaReferences

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {356B06EC-4908-42A4-81FC-4B5A51F3483B}  

A collection of XMLSchemaReference objects that represent the unique namespaces that are attached to a document.

**Remarks:** Use the XMLSchemaReferences property to return a collection of schemas attached to a document. The following example loops through the schemas attached to a document. If it finds the specified schema, it reloads it; if it doesn't find the specified schema, it attaches the schema to the document.

## Properties (8)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of XML schemas in the collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified XMLSchemaReferences object.
- `HideValidationErrors As Boolean  (read/write)`  
  Returns a Boolean indicating whether Word displays schema validation errors for the current XML document. Read/write.
- `IgnoreMixedContent As Boolean  (read/write)`  
  Returns a Boolean that represents whether Microsoft Word preforms validation on text nodes that have element siblings and specifies whether these text nodes are saved in XML when the XMLSaveDataOnly property is True. Read/write.
- `ShowPlaceholderText As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether automatic placeholder text is displayed for XML elements in a document. Read/write.

## Methods (3)

- `Item(Index As Variant) As XMLSchemaReference`  
  Returns an individual XMLSchemaReference object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `Validate()`  
  Validates all the XML schemas that are attached to a document.
- `Add([NamespaceURI As Variant], [Alias As Variant], [FileName As Variant], [InstallForAllUsers As Boolean]) As XMLSchemaReference`  
  Returns an XMLSchemaReference that represents a schema applied to a document.
    - `NamespaceURI As Variant` (optional): The name of the schema as defined in the schema. The Namespace parameter is case-sensitive and must be spelled exactly as it appears in the schema. If the specified namespace cannot be found in any of the schemas attached to the document, an error is displayed.
    - `Alias As Variant` (optional): The name of the schema as it appears on the Schemas tab in the Templates and Add-ins dialog box.
    - `FileName As Variant` (optional): The path and file name of the schema. This may be a local file path, a network path, or an Internet address.
    - `InstallForAllUsers As Boolean` (optional): True if all users that log on to a computer can access and use the new schema. The default is False.
