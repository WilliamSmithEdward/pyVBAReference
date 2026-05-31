# IOLEObjects

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208A3-0001-0000-C000-000000000046}  

## Properties (20)

- `Application As HRESULT  (read-only)`
- `Creator As HRESULT  (read-only)`
- `Parent As HRESULT  (read-only)`
- `Enabled As HRESULT  (read/write)`
- `Height As HRESULT  (read/write)`
- `Left As HRESULT  (read/write)`
- `Locked As HRESULT  (read/write)`
- `Placement As HRESULT  (read/write)`
- `PrintObject As HRESULT  (read/write)`
- `Top As HRESULT  (read/write)`
- `Visible As HRESULT  (read/write)`
- `Width As HRESULT  (read/write)`
- `ZOrder As HRESULT  (read-only)`
- `ShapeRange As HRESULT  (read-only)`
- `Border As HRESULT  (read-only)`
- `Interior As HRESULT  (read-only)`
- `Shadow As HRESULT  (read/write)`
- `AutoLoad As HRESULT  (read/write)`
- `SourceName As HRESULT  (read/write)`
- `Count As HRESULT  (read-only)`

## Methods (11)

- `BringToFront(RHS As Variant)`
- `Copy(RHS As Variant)`
- `CopyPicture([Appearance As XlPictureAppearance], [Format As XlCopyPictureFormat], RHS As Variant)`
- `Cut(RHS As Variant)`
- `Delete(RHS As Variant)`
- `Duplicate(RHS As Object)`
- `Select([Replace As Variant], RHS As Variant)`
- `SendToBack(RHS As Variant)`
- `Add([ClassType As Variant], [Filename As Variant], [Link As Variant], [DisplayAsIcon As Variant], [IconFileName As Variant], [IconIndex As Variant], [IconLabel As Variant], [Left As Variant], [Top As Variant], [Width As Variant], [Height As Variant], RHS As OLEObject)`
- `Item(Index As Variant, RHS As Object)`
- `_NewEnum(RHS As IUnknown)`
