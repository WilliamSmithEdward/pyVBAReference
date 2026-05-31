# IOLEObject

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208A2-0001-0000-C000-000000000046}  

## Properties (29)

- `Application As HRESULT  (read-only)`
- `Creator As HRESULT  (read-only)`
- `Parent As HRESULT  (read-only)`
- `BottomRightCell As HRESULT  (read-only)`
- `Enabled As HRESULT  (read/write)`
- `Height As HRESULT  (read/write)`
- `Index As HRESULT  (read-only)`
- `Left As HRESULT  (read/write)`
- `Locked As HRESULT  (read/write)`
- `Name As HRESULT  (read/write)`
- `Placement As HRESULT  (read/write)`
- `PrintObject As HRESULT  (read/write)`
- `Top As HRESULT  (read/write)`
- `TopLeftCell As HRESULT  (read-only)`
- `Visible As HRESULT  (read/write)`
- `Width As HRESULT  (read/write)`
- `ZOrder As HRESULT  (read-only)`
- `ShapeRange As HRESULT  (read-only)`
- `Border As HRESULT  (read-only)`
- `Interior As HRESULT  (read-only)`
- `Shadow As HRESULT  (read/write)`
- `AutoLoad As HRESULT  (read/write)`
- `AutoUpdate As HRESULT  (read/write)`
- `Object As HRESULT  (read-only)`
- `OLEType As HRESULT  (read-only)`
- `SourceName As HRESULT  (read/write)`
- `LinkedCell As HRESULT  (read/write)`
- `ListFillRange As HRESULT  (read/write)`
- `progID As HRESULT  (read-only)`

## Methods (11)

- `BringToFront(RHS As Variant)`
- `Copy(RHS As Variant)`
- `CopyPicture([Appearance As XlPictureAppearance], [Format As XlCopyPictureFormat], RHS As Variant)`
- `Cut(RHS As Variant)`
- `Delete(RHS As Variant)`
- `Duplicate(RHS As Object)`
- `Select([Replace As Variant], RHS As Variant)`
- `SendToBack(RHS As Variant)`
- `Activate(RHS As Variant)`
- `Update(RHS As Variant)`
- `Verb([Verb As XlOLEVerb], RHS As Variant)`
