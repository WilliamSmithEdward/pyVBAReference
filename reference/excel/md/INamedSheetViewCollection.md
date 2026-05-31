# INamedSheetViewCollection

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024501-0001-0000-C000-000000000046}  

## Properties (5)

- `Application As HRESULT  (read-only)`
- `Creator As HRESULT  (read-only)`
- `Parent As HRESULT  (read-only)`
- `Count As HRESULT  (read-only)`
- `_NewEnum As HRESULT  (read-only)`

## Methods (6)

- `Add(Name As String, RHS As NamedSheetView)`
- `EnterTemporary(RHS As NamedSheetView)`
- `Exit()`
- `GetActive(RHS As NamedSheetView)`
- `GetItem(Name As String, RHS As NamedSheetView)`
- `GetItemAt(Index As Long, RHS As NamedSheetView)`
