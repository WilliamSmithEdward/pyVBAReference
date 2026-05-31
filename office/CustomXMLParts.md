# CustomXMLParts

**Type:** Class  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000CDB0C-0000-0000-C000-000000000046}  

## Properties (6)

- `Application As Object  (read-only)`
- `Creator As Long  (read-only)`
- `Parent As Object  (read-only)`
- `Count As Long  (read-only)`
- `Item As CustomXMLPart  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (3)

- `Add([XML As String], [SchemaCollection As Variant]) As CustomXMLPart`
- `SelectByID(Id As String) As CustomXMLPart`
- `SelectByNamespace(NamespaceURI As String) As CustomXMLParts`

## Events (3)

- `PartAfterAdd(NewPart As CustomXMLPart)`
- `PartBeforeDelete(OldPart As CustomXMLPart)`
- `PartAfterLoad(Part As CustomXMLPart)`
