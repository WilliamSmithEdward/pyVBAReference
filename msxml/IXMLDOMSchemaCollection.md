# IXMLDOMSchemaCollection

**Type:** Dispatch Interface  
**Library:** Microsoft XML, v6.0  
**GUID:** {373984C8-B845-449B-91E7-45AC83036ADE}  

XML Schemas Collection

## Properties (2)

- `length As Long  (read-only)`  
  number of schemas in collection
- `namespaceURI As String  (read-only)`  
  Get namespaceURI for schema by index

## Methods (4)

- `add(namespaceURI As String, var As Variant)`  
  add a new schema
- `get(namespaceURI As String) As IXMLDOMNode`  
  lookup schema by namespaceURI
- `remove(namespaceURI As String)`  
  remove schema by namespaceURI
- `addCollection(otherCollection As IXMLDOMSchemaCollection)`  
  copy & merge other collection into this one
