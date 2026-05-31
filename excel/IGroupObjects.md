# IGroupObjects

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020899-0001-0000-C000-000000000046}  

## Properties (30)

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
- `AddIndent As HRESULT  (read/write)`
- `ArrowHeadLength As HRESULT  (read/write)`
- `ArrowHeadStyle As HRESULT  (read/write)`
- `ArrowHeadWidth As HRESULT  (read/write)`
- `AutoSize As HRESULT  (read/write)`
- `Border As HRESULT  (read-only)`
- `_Default As HRESULT  (read/write)`
- `Font As HRESULT  (read-only)`
- `HorizontalAlignment As HRESULT  (read/write)`
- `Interior As HRESULT  (read-only)`
- `Orientation As HRESULT  (read/write)`
- `RoundedCorners As HRESULT  (read/write)`
- `Shadow As HRESULT  (read/write)`
- `VerticalAlignment As HRESULT  (read/write)`
- `ReadingOrder As HRESULT  (read/write)`
- `Count As HRESULT  (read-only)`

## Methods (13)

- `BringToFront(RHS As Variant)`
- `Copy(RHS As Variant)`
- `CopyPicture([Appearance As XlPictureAppearance], [Format As XlCopyPictureFormat], RHS As Variant)`
- `Cut(RHS As Variant)`
- `Delete(RHS As Variant)`
- `Duplicate(RHS As Object)`
- `Select([Replace As Variant], RHS As Variant)`
- `SendToBack(RHS As Variant)`
- `CheckSpelling([CustomDictionary As Variant], [IgnoreUppercase As Variant], [AlwaysSuggest As Variant], [SpellLang As Variant], RHS As Variant)`
- `Ungroup(RHS As Object)`
- `Group(RHS As GroupObject)`
- `Item(Index As Variant, RHS As Object)`
- `_NewEnum(RHS As IUnknown)`
