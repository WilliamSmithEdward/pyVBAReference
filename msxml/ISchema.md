# ISchema

**Type:** Dispatch Interface  
**Library:** Microsoft XML, v6.0  
**GUID:** {50EA08B4-DD1B-4664-9A50-C2F40F4BD79A}  

XML Schema

## Properties (15)

- `name As String  (read-only)`
- `namespaceURI As String  (read-only)`
- `schema As ISchema  (read-only)`
- `id As String  (read-only)`
- `itemType As SOMITEMTYPE  (read-only)`
- `unhandledAttributes As IVBSAXAttributes  (read-only)`
- `targetNamespace As String  (read-only)`
- `version As String  (read-only)`
- `types As ISchemaItemCollection  (read-only)`
- `elements As ISchemaItemCollection  (read-only)`
- `attributes As ISchemaItemCollection  (read-only)`
- `attributeGroups As ISchemaItemCollection  (read-only)`
- `modelGroups As ISchemaItemCollection  (read-only)`
- `notations As ISchemaItemCollection  (read-only)`
- `schemaLocations As ISchemaStringCollection  (read-only)`

## Methods (1)

- `writeAnnotation(annotationSink As IUnknown) As Boolean`
