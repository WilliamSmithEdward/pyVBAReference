# IDrawingObjects

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002086F-0001-0000-C000-000000000046}  

## Properties (54)

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
- `Accelerator As HRESULT  (read/write)`
- `AddIndent As HRESULT  (read/write)`
- `ArrowHeadLength As HRESULT  (read/write)`
- `ArrowHeadStyle As HRESULT  (read/write)`
- `ArrowHeadWidth As HRESULT  (read/write)`
- `AutoSize As HRESULT  (read/write)`
- `Border As HRESULT  (read-only)`
- `CancelButton As HRESULT  (read/write)`
- `Caption As HRESULT  (read/write)`
- `Characters As HRESULT  (read-only)`
- `_Default As HRESULT  (read/write)`
- `DefaultButton As HRESULT  (read/write)`
- `DismissButton As HRESULT  (read/write)`
- `Display3DShading As HRESULT  (read/write)`
- `DisplayVerticalScrollBar As HRESULT  (read/write)`
- `DropDownLines As HRESULT  (read/write)`
- `Font As HRESULT  (read-only)`
- `HelpButton As HRESULT  (read/write)`
- `HorizontalAlignment As HRESULT  (read/write)`
- `InputType As HRESULT  (read/write)`
- `Interior As HRESULT  (read-only)`
- `LargeChange As HRESULT  (read/write)`
- `LinkedCell As HRESULT  (read/write)`
- `ListFillRange As HRESULT  (read/write)`
- `ListIndex As HRESULT  (read/write)`
- `LockedText As HRESULT  (read/write)`
- `Max As HRESULT  (read/write)`
- `Min As HRESULT  (read/write)`
- `MultiLine As HRESULT  (read/write)`
- `MultiSelect As HRESULT  (read/write)`
- `Orientation As HRESULT  (read/write)`
- `PhoneticAccelerator As HRESULT  (read/write)`
- `RoundedCorners As HRESULT  (read/write)`
- `Shadow As HRESULT  (read/write)`
- `SmallChange As HRESULT  (read/write)`
- `Text As HRESULT  (read/write)`
- `Value As HRESULT  (read/write)`
- `VerticalAlignment As HRESULT  (read/write)`
- `ReadingOrder As HRESULT  (read/write)`
- `Count As HRESULT  (read-only)`

## Methods (21)

- `BringToFront(RHS As Variant)`
- `Copy(RHS As Variant)`
- `CopyPicture([Appearance As XlPictureAppearance], [Format As XlCopyPictureFormat], RHS As Variant)`
- `Cut(RHS As Variant)`
- `Delete(RHS As Variant)`
- `Duplicate(RHS As Object)`
- `Select([Replace As Variant], RHS As Variant)`
- `SendToBack(RHS As Variant)`
- `AddItem(Text As Variant, [Index As Variant], RHS As Variant)`
- `CheckSpelling([CustomDictionary As Variant], [IgnoreUppercase As Variant], [AlwaysSuggest As Variant], [SpellLang As Variant], RHS As Variant)`
- `List([Index As Variant], RHS As Variant)`
- `RemoveAllItems(RHS As Variant)`
- `RemoveItem(Index As Long, [Count As Variant], RHS As Variant)`
- `Reshape(Vertex As Long, Insert As Variant, [Left As Variant], [Top As Variant], RHS As Variant)`
- `Selected([Index As Variant], RHS As Variant)`
- `Ungroup(RHS As Object)`
- `Vertices([Index1 As Variant], [Index2 As Variant], RHS As Variant)`
- `Item(Index As Variant, RHS As Object)`
- `Group(RHS As GroupObject)`
- `LinkCombo([Link As Variant], RHS As Variant)`
- `_NewEnum(RHS As IUnknown)`
