# IEditBoxes

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020884-0001-0000-C000-000000000046}  

## Properties (23)

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
- `Caption As HRESULT  (read/write)`
- `Characters As HRESULT  (read-only)`
- `LockedText As HRESULT  (read/write)`
- `Text As HRESULT  (read/write)`
- `DisplayVerticalScrollBar As HRESULT  (read/write)`
- `InputType As HRESULT  (read/write)`
- `MultiLine As HRESULT  (read/write)`
- `PasswordEdit As HRESULT  (read/write)`
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
- `Add(Left As Double, Top As Double, Width As Double, Height As Double, RHS As EditBox)`
- `Group(RHS As GroupObject)`
- `Item(Index As Variant, RHS As Variant)`
- `_NewEnum(RHS As IUnknown)`
