# FileDialogFilters

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0365-0000-0000-C000-000000000046}  

## Properties (5)

- `Application As Object  (read-only)`
- `Creator As Long  (read-only)`
- `Parent As Object  (read-only)`
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`

## Methods (4)

- `Item(Index As Long) As FileDialogFilter`
- `Delete([filter As Variant])`
- `Clear()`
- `Add(Description As String, Extensions As String, [Position As Variant]) As FileDialogFilter`
