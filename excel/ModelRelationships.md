# ModelRelationships

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244DA-0000-0000-C000-000000000046}  

## Properties (6)

- `Application As Application  (read-only)`
- `Creator As XlCreator  (read-only)`
- `Parent As Object  (read-only)`
- `Count As Long  (read-only)`
- `_Default As ModelRelationship  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (3)

- `Item(Index As Variant) As ModelRelationship`
- `Add(ForeignKeyColumn As ModelTableColumn, PrimaryKeyColumn As ModelTableColumn) As ModelRelationship`
- `DetectRelationships(PivotTable As PivotTable)`
