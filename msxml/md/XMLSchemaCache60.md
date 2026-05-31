# XMLSchemaCache60

**Type:** Class  
**Library:** Microsoft XML, v6.0  
**GUID:** {88D96A07-F192-11D4-A65F-0040963251E5}  

XML Schema Cache 6.0

## Properties (3)

- `length As Long  (read-only)`  
  number of schemas in collection
- `namespaceURI As String  (read-only)`  
  Get namespaceURI for schema by index
- `validateOnLoad As Boolean  (read/write)`

## Methods (7)

- `add(namespaceURI As String, var As Variant)`  
  add a new schema
- `get(namespaceURI As String) As IXMLDOMNode`  
  lookup schema by namespaceURI
- `remove(namespaceURI As String)`  
  remove schema by namespaceURI
- `addCollection(otherCollection As IXMLDOMSchemaCollection)`  
  copy & merge other collection into this one
- `validate()`
- `getSchema(namespaceURI As String) As ISchema`
- `getDeclaration(node As IXMLDOMNode) As ISchemaItem`
