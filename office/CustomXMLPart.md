# CustomXMLPart

**Type:** Class  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000CDB08-0000-0000-C000-000000000046}  

## Properties (11)

- `Application As Object  (read-only)`
- `Creator As Long  (read-only)`
- `Parent As Object  (read-only)`
- `DocumentElement As CustomXMLNode  (read-only)`
- `Id As String  (read-only)`
- `NamespaceURI As String  (read-only)`
- `SchemaCollection As CustomXMLSchemaCollection  (read/write)`
- `NamespaceManager As CustomXMLPrefixMappings  (read-only)`
- `XML As String  (read-only)`
- `Errors As CustomXMLValidationErrors  (read-only)`
- `BuiltIn As Boolean  (read-only)`

## Methods (6)

- `AddNode(Parent As CustomXMLNode, [Name As String], [NamespaceURI As String], [NextSibling As CustomXMLNode], [NodeType As MsoCustomXMLNodeType], [NodeValue As String])`
- `Delete()`
- `Load(FilePath As String) As Boolean`
- `LoadXML(XML As String) As Boolean`
- `SelectNodes(XPath As String) As CustomXMLNodes`
- `SelectSingleNode(XPath As String) As CustomXMLNode`

## Events (3)

- `NodeAfterInsert(NewNode As CustomXMLNode, InUndoRedo As Boolean)`
- `NodeAfterDelete(OldNode As CustomXMLNode, OldParentNode As CustomXMLNode, OldNextSibling As CustomXMLNode, InUndoRedo As Boolean)`
- `NodeAfterReplace(OldNode As CustomXMLNode, NewNode As CustomXMLNode, InUndoRedo As Boolean)`
