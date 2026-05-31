# IModule

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208AD-0001-0000-C000-000000000046}  

## Properties (14)

- `Application As HRESULT  (read-only)`
- `Creator As HRESULT  (read-only)`
- `Parent As HRESULT  (read-only)`
- `CodeName As HRESULT  (read-only)`
- `_CodeName As HRESULT  (read/write)`
- `Index As HRESULT  (read-only)`
- `Name As HRESULT  (read/write)`
- `Next As HRESULT  (read-only)`
- `PageSetup As HRESULT  (read-only)`
- `Previous As HRESULT  (read-only)`
- `ProtectContents As HRESULT  (read-only)`
- `ProtectionMode As HRESULT  (read-only)`
- `Visible As HRESULT  (read/write)`
- `Shapes As HRESULT  (read-only)`

## Methods (10)

- `Activate(lcid As Long)`
- `Copy([Before As Variant], [After As Variant], lcid As Long)`
- `Delete(lcid As Long)`
- `Move([Before As Variant], [After As Variant], lcid As Long)`
- `Select([Replace As Variant], lcid As Long)`
- `Unprotect([Password As Variant], lcid As Long)`
- `InsertFile(Filename As Variant, [Merge As Variant], RHS As Variant)`
- `Protect([Password As Variant], [DrawingObjects As Variant], [Contents As Variant], [Scenarios As Variant], [UserInterfaceOnly As Variant])`
- `PrintOut([From As Variant], [To As Variant], [Copies As Variant], [Preview As Variant], [ActivePrinter As Variant], [PrintToFile As Variant], [Collate As Variant])`
- `SaveAs(Filename As String, [FileFormat As Variant], [Password As Variant], [WriteResPassword As Variant], [ReadOnlyRecommended As Variant], [CreateBackup As Variant], [AddToMru As Variant], [TextCodepage As Variant], [TextVisualLayout As Variant])`
