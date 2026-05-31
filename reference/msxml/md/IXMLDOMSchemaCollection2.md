# IXMLDOMSchemaCollection2

**Type:** Dispatch Interface  
**Library:** Microsoft XML, v6.0  
**GUID:** {50EA08B0-DD1B-4664-9A50-C2F40F4BD79A}  

XML Schemas Collection 2

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
