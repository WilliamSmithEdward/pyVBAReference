# Pages

**Type:** Dispatch Interface  
**Library:** Microsoft Forms 2.0 Object Library  
**GUID:** {92E11A03-7358-11CE-80CB-00AA00611080}  

## Properties (2)

- `Count As Long  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (9)

- `Item(varg As Variant) As Object`
- `Enum() As IUnknown`
- `Add([bstrName As Variant], [bstrCaption As Variant], [lIndex As Variant]) As Page`
- `_AddCtrl(clsid As Long, bstrName As String, bstrCaption As String) As Page`
- `_InsertCtrl(clsid As Long, bstrName As String, bstrCaption As String, lIndex As Long) As Page`
- `_GetItemByIndex(lIndex As Long) As Control`
- `_GetItemByName(pstrName As String) As Control`
- `Remove(varg As Variant)`
- `Clear()`
