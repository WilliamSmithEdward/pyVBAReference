# Connections

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024486-0000-0000-C000-000000000046}  

## Properties (6)

- `Application As Application  (read-only)`
- `Creator As XlCreator  (read-only)`
- `Parent As Object  (read-only)`
- `Count As Long  (read-only)`
- `_Default As WorkbookConnection  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (3)

- `Item(Index As Variant) As WorkbookConnection`
- `Add2(Name As String, Description As String, ConnectionString As Variant, CommandText As Variant, [lCmdtype As Variant], [CreateModelConnection As Variant], [ImportRelationships As Variant]) As WorkbookConnection`
- `AddFromFile(Filename As String, [CreateModelConnection As Variant], [ImportRelationships As Variant]) As WorkbookConnection`
