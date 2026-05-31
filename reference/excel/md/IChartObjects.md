# IChartObjects

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208D0-0001-0000-C000-000000000046}  

## Properties (14)

- `Application As HRESULT  (read-only)`
- `Creator As HRESULT  (read-only)`
- `Parent As HRESULT  (read-only)`
- `Height As HRESULT  (read/write)`
- `Left As HRESULT  (read/write)`
- `Locked As HRESULT  (read/write)`
- `Placement As HRESULT  (read/write)`
- `PrintObject As HRESULT  (read/write)`
- `Top As HRESULT  (read/write)`
- `Visible As HRESULT  (read/write)`
- `Width As HRESULT  (read/write)`
- `ShapeRange As HRESULT  (read-only)`
- `ProtectChartObject As HRESULT  (read/write)`
- `Count As HRESULT  (read-only)`

## Methods (10)

- `CopyPicture([Appearance As XlPictureAppearance], [Format As XlCopyPictureFormat], RHS As Variant)`
- `Cut(RHS As Variant)`
- `Delete(RHS As Variant)`
- `Duplicate(RHS As Object)`
- `Select([Replace As Variant], RHS As Variant)`
- `Copy(RHS As Variant)`
- `Add(Left As Double, Top As Double, Width As Double, Height As Double, RHS As ChartObject)`
- `Item(Index As Variant, RHS As Object)`
- `_NewEnum(RHS As IUnknown)`
- `_Default(Index As Variant, RHS As Object)`
