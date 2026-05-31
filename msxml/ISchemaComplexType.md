# ISchemaComplexType

**Type:** Dispatch Interface  
**Library:** Microsoft XML, v6.0  
**GUID:** {50EA08B9-DD1B-4664-9A50-C2F40F4BD79A}  

XML Schema Complex Type

## Properties (28)

- `name As String  (read-only)`
- `namespaceURI As String  (read-only)`
- `schema As ISchema  (read-only)`
- `id As String  (read-only)`
- `itemType As SOMITEMTYPE  (read-only)`
- `unhandledAttributes As IVBSAXAttributes  (read-only)`
- `baseTypes As ISchemaItemCollection  (read-only)`
- `final As SCHEMADERIVATIONMETHOD  (read-only)`
- `variety As SCHEMATYPEVARIETY  (read-only)`
- `derivedBy As SCHEMADERIVATIONMETHOD  (read-only)`
- `minExclusive As String  (read-only)`
- `minInclusive As String  (read-only)`
- `maxExclusive As String  (read-only)`
- `maxInclusive As String  (read-only)`
- `totalDigits As Variant  (read-only)`
- `fractionDigits As Variant  (read-only)`
- `length As Variant  (read-only)`
- `minLength As Variant  (read-only)`
- `maxLength As Variant  (read-only)`
- `enumeration As ISchemaStringCollection  (read-only)`
- `whitespace As SCHEMAWHITESPACE  (read-only)`
- `patterns As ISchemaStringCollection  (read-only)`
- `isAbstract As Boolean  (read-only)`
- `anyAttribute As ISchemaAny  (read-only)`
- `attributes As ISchemaItemCollection  (read-only)`
- `contentType As SCHEMACONTENTTYPE  (read-only)`
- `contentModel As ISchemaModelGroup  (read-only)`
- `prohibitedSubstitutions As SCHEMADERIVATIONMETHOD  (read-only)`

## Methods (2)

- `writeAnnotation(annotationSink As IUnknown) As Boolean`
- `isValid(data As String) As Boolean`
