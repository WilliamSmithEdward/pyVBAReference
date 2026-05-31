# IFont

**Type:** Interface  
**Library:** Microsoft Forms 2.0 Object Library  
**GUID:** {BEF6E002-A874-101A-8BBA-00AA00300CAB}  

## Properties (9)

- `Name As HRESULT  (read/write)`
- `Size As HRESULT  (read/write)`
- `Bold As HRESULT  (read/write)`
- `Italic As HRESULT  (read/write)`
- `Underline As HRESULT  (read/write)`
- `Strikethrough As HRESULT  (read/write)`
- `Weight As HRESULT  (read/write)`
- `Charset As HRESULT  (read/write)`
- `hFont As HRESULT  (read-only)`

## Methods (5)

- `Clone(lplpfont As IFont)`
- `IsEqual(lpFontOther As IFont)`
- `SetRatio(cyLogical As Long, cyHimetric As Long)`
- `AddRefHfont(hFont As OLE_HANDLE)`
- `ReleaseHfont(hFont As OLE_HANDLE)`
