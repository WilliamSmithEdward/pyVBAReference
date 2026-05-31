# ISchemaElement

**Type:** Dispatch Interface  
**Library:** Microsoft XML, v6.0  
**GUID:** {50EA08B7-DD1B-4664-9A50-C2F40F4BD79A}  

XML Schema Element

## Properties (19)

- `name As String  (read-only)`
- `namespaceURI As String  (read-only)`
- `schema As ISchema  (read-only)`
- `id As String  (read-only)`
- `itemType As SOMITEMTYPE  (read-only)`
- `unhandledAttributes As IVBSAXAttributes  (read-only)`
- `minOccurs As Variant  (read-only)`
- `maxOccurs As Variant  (read-only)`
- `type As ISchemaType  (read-only)`
- `scope As ISchemaComplexType  (read-only)`
- `defaultValue As String  (read-only)`
- `fixedValue As String  (read-only)`
- `isNillable As Boolean  (read-only)`
- `identityConstraints As ISchemaItemCollection  (read-only)`
- `substitutionGroup As ISchemaElement  (read-only)`
- `substitutionGroupExclusions As SCHEMADERIVATIONMETHOD  (read-only)`
- `disallowedSubstitutions As SCHEMADERIVATIONMETHOD  (read-only)`
- `isAbstract As Boolean  (read-only)`
- `isReference As Boolean  (read-only)`

## Methods (1)

- `writeAnnotation(annotationSink As IUnknown) As Boolean`
