# Windows

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020892-0000-0000-C000-000000000046}  

## Properties (8)

- `Application As Application  (read-only)`
- `Creator As XlCreator  (read-only)`
- `Parent As Object  (read-only)`
- `Count As Long  (read-only)`
- `Item As Window  (read-only)`
- `_NewEnum As IUnknown  (read-only)`
- `_Default As Window  (read-only)`
- `SyncScrollingSideBySide As Boolean  (read/write)`

## Methods (4)

- `Arrange([ArrangeStyle As XlArrangeStyle], [ActiveWorkbook As Variant], [SyncHorizontal As Variant], [SyncVertical As Variant]) As Variant`
- `CompareSideBySideWith(WindowName As Variant) As Boolean`
- `BreakSideBySide() As Boolean`
- `ResetPositionsSideBySide()`
