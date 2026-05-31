# IDialogSheets

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208B0-0001-0000-C000-000000000046}  

## Properties (10)

- `Application As HRESULT  (read-only)`
- `Creator As HRESULT  (read-only)`
- `Parent As HRESULT  (read-only)`
- `Count As HRESULT  (read-only)`
- `Item As HRESULT  (read-only)`
- `_NewEnum As HRESULT  (read-only)`
- `HPageBreaks As HRESULT  (read-only)`
- `VPageBreaks As HRESULT  (read-only)`
- `Visible As HRESULT  (read/write)`
- `_Default As HRESULT  (read-only)`

## Methods (8)

- `Add([Before As Variant], [After As Variant], [Count As Variant], RHS As DialogSheet)`
- `Copy([Before As Variant], [After As Variant], lcid As Long)`
- `Delete(lcid As Long)`
- `Move([Before As Variant], [After As Variant], lcid As Long)`
- `PrintPreview([EnableChanges As Variant], lcid As Long)`
- `Select([Replace As Variant], lcid As Long)`
- `PrintOut([From As Variant], [To As Variant], [Copies As Variant], [Preview As Variant], [ActivePrinter As Variant], [PrintToFile As Variant], [Collate As Variant], [PrToFileName As Variant])`
- `Add2([Before As Variant], [After As Variant], [Count As Variant], [NewLayout As Variant], RHS As Object)`
