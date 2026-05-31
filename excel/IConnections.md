# IConnections

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024486-0001-0000-C000-000000000046}  

## Properties (6)

- `Application As HRESULT  (read-only)`
- `Creator As HRESULT  (read-only)`
- `Parent As HRESULT  (read-only)`
- `Count As HRESULT  (read-only)`
- `_Default As HRESULT  (read-only)`
- `_NewEnum As HRESULT  (read-only)`

## Methods (3)

- `Item(Index As Variant, RHS As WorkbookConnection)`
- `Add2(Name As String, Description As String, ConnectionString As Variant, CommandText As Variant, [lCmdtype As Variant], [CreateModelConnection As Variant], [ImportRelationships As Variant], RHS As WorkbookConnection)`
- `AddFromFile(Filename As String, [CreateModelConnection As Variant], [ImportRelationships As Variant], RHS As WorkbookConnection)`
