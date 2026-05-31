# IFont

**Type:** Interface  
**Library:** OLE Automation  
**GUID:** {BEF6E002-A874-101A-8BBA-00AA00300CAB}  

Font Object

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

- `Clone(ppfont As IFont)`
- `IsEqual(pfontOther As IFont)`
- `SetRatio(cyLogical As Long, cyHimetric As Long)`
- `AddRefHfont(hFont As OLE_HANDLE)`
- `ReleaseHfont(hFont As OLE_HANDLE)`
